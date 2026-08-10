"""Génération des rapports HTML autonomes de POPS Check.

Le module ne dépend d'aucun moteur de templates et accepte aussi bien des
dataclasses/objets que des dictionnaires. Toutes les valeurs provenant des
résultats sont échappées avant d'être injectées dans le HTML.

API publique principale::

    index = generate_reports(results, "rapports", run_metadata={...})

``index`` est le :class:`~pathlib.Path` du rapport global. Un rapport détaillé
est créé à côté pour chaque pays.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from hashlib import sha256
from html import escape
import json
from pathlib import Path
import re
import unicodedata
from typing import Any, Iterable, Mapping, MutableSet, Optional, Sequence


__all__ = [
    "ReportGenerator",
    "generate_country_report",
    "generate_index_report",
    "generate_reports",
    "render_country_html",
    "render_index_html",
    "safe_slug",
    "slugify_country",
]


_UNSET = object()
_RESULT_FIELDS = {
    "key",
    "country",
    "pays",
    "reference_path",
    "reference_file",
    "received_path",
    "received_file",
    "status",
    "anomalies",
    "issues",
    "errors",
}
_CATEGORY_ORDER = ("feuilles", "colonnes", "lignes", "fichier", "autres")
_CATEGORY_LABELS = {
    "feuilles": "Feuilles",
    "colonnes": "Colonnes",
    "lignes": "Lignes",
    "fichier": "Fichier",
    "autres": "Autres",
}


@dataclass(frozen=True)
class _Issue:
    category: str
    code: str
    message: str
    sheet: Any = None
    element: Any = None
    expected_position: Any = None
    observed_position: Any = None
    expected: Any = None
    observed: Any = None
    severity: str = "warning"
    impact: int = 1
    details: Mapping[str, Any] = field(default_factory=dict)


@dataclass
class _Country:
    key: str
    country: str
    reference_path: Any
    received_path: Any
    status_code: str
    status_label: str
    tone: str
    anomalies: list[_Issue]
    errors: list[str]
    warnings: list[str]
    sheet_order_expected: list[Any]
    sheet_order_observed: list[Any]
    metadata: Mapping[str, Any]
    counts: dict[str, int]
    total_anomalies: int
    missing_reason: str = ""
    filename: str = ""


def _text(value: Any, default: str = "") -> str:
    """Convertit une valeur quelconque en texte sans propager ``__str__``."""

    if value is None:
        return default
    try:
        if isinstance(value, Enum):
            value = value.value
        if isinstance(value, Path):
            return str(value)
        if isinstance(value, bool):
            return "Oui" if value else "Non"
        if isinstance(value, Mapping):
            return json.dumps(value, ensure_ascii=False, default=str, sort_keys=True)
        if isinstance(value, (list, tuple, set, frozenset)):
            return " › ".join(_text(item, "—") for item in value)
        return str(value)
    except Exception:
        return default


def _html(value: Any, default: str = "—") -> str:
    rendered = _text(value, default)
    if not rendered:
        rendered = default
    return escape(rendered, quote=True)


def _token(value: Any) -> str:
    raw = _text(value).strip().casefold()
    raw = unicodedata.normalize("NFKD", raw)
    raw = "".join(char for char in raw if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "_", raw).strip("_")


def _mapping_lookup(mapping: Mapping[Any, Any], name: str) -> tuple[bool, Any]:
    if name in mapping:
        return True, mapping[name]
    wanted = _token(name)
    for key, value in mapping.items():
        if _token(key) == wanted:
            return True, value
    return False, None


def _lookup(obj: Any, name: str) -> tuple[bool, Any]:
    current = obj
    for part in name.split("."):
        if current is None:
            return False, None
        if isinstance(current, Mapping):
            found, current = _mapping_lookup(current, part)
            if not found:
                return False, None
            continue
        try:
            current = getattr(current, part)
        except (AttributeError, TypeError, ValueError):
            return False, None
        except Exception:
            return False, None
    return True, current


def _get(obj: Any, *names: str, default: Any = None) -> Any:
    for name in names:
        found, value = _lookup(obj, name)
        if found:
            return value
    return default


def _has(obj: Any, *names: str) -> bool:
    return any(_lookup(obj, name)[0] for name in names)


def _list(value: Any) -> list[Any]:
    if value is None or value is False:
        return []
    if isinstance(value, (str, bytes, bytearray, Path, BaseException)):
        return [value]
    if isinstance(value, Mapping):
        return [value]
    try:
        return list(value)
    except (TypeError, ValueError):
        return [value]


def _integer(value: Any) -> Optional[int]:
    if value is None or isinstance(value, bool):
        return None
    try:
        number = int(value)
    except (TypeError, ValueError, OverflowError):
        return None
    return max(0, number)


def _category(value: Any, code: Any = "", message: Any = "") -> str:
    explicit = _token(value)
    explicit_aliases = {
        "feuilles": {"feuille", "feuilles", "sheet", "sheets", "worksheet", "worksheets"},
        "colonnes": {"colonne", "colonnes", "column", "columns", "col"},
        "lignes": {"ligne", "lignes", "row", "rows"},
        "fichier": {"fichier", "fichiers", "file", "files", "workbook", "classeur"},
        "autres": {"autre", "autres", "other", "others"},
    }
    for category, aliases in explicit_aliases.items():
        if explicit in aliases:
            return category
    code_token = _token(code)
    if code_token.startswith(("sheet_", "worksheet_")):
        return "feuilles"
    if code_token.startswith(("column_", "col_")):
        return "colonnes"
    if code_token.startswith("row_"):
        return "lignes"
    if code_token.startswith(("file_", "workbook_", "reference_", "received_", "invalid_")):
        return "fichier"
    # Le message sert uniquement de repli : il contient souvent « dans la
    # feuille X » même lorsque la catégorie réelle est une ligne ou colonne.
    haystack = "_".join((code_token, _token(message)))
    if any(word in haystack for word in ("feuille", "sheet", "worksheet", "onglet", "tab_")):
        return "feuilles"
    if any(word in haystack for word in ("colonne", "column", "_col_")):
        return "colonnes"
    if any(word in haystack for word in ("ligne", "row")):
        return "lignes"
    if any(
        word in haystack
        for word in (
            "fichier",
            "file",
            "workbook",
            "classeur",
            "reference_missing",
            "received_missing",
            "invalid",
        )
    ):
        return "fichier"
    return "autres"


def _normalise_details(value: Any) -> Mapping[str, Any]:
    if isinstance(value, Mapping):
        return value
    return {}


def _detail_value(details: Mapping[str, Any], *names: str) -> Any:
    return _get(details, *names, default=None)


def _normalise_issue(raw: Any, forced_category: str = "") -> _Issue:
    if isinstance(raw, (str, bytes, bytearray, BaseException)):
        message = _text(raw, "Anomalie non détaillée")
        return _Issue(
            category=forced_category or _category("", "", message),
            code="ANOMALY",
            message=message,
        )

    details = _normalise_details(_get(raw, "details", "detail", "evidence", default={}))
    code = _text(
        _get(raw, "code", "anomaly_code", "type", "kind", "type_anomalie", default="ANOMALY"),
        "ANOMALY",
    )
    message = _text(
        _get(raw, "message", "description", "label", "detail_message", default=""),
        "",
    )
    if not message:
        message = code.replace("_", " ").strip().capitalize() or "Anomalie non détaillée"
    raw_category = _get(raw, "category", "categorie", "domain", "scope", default="")
    category = forced_category or _category(raw_category, code, message)

    expected_position = _get(
        raw,
        "expected_position",
        "position_expected",
        "position_attendue",
        "expected_index",
        default=_UNSET,
    )
    if expected_position is _UNSET:
        expected_position = _detail_value(
            details, "expected_position", "position_expected", "position_attendue", "expected_index"
        )
    observed_position = _get(
        raw,
        "observed_position",
        "actual_position",
        "received_position",
        "position_observed",
        "position_recue",
        "observed_index",
        "actual_index",
        default=_UNSET,
    )
    if observed_position is _UNSET:
        observed_position = _detail_value(
            details,
            "observed_position",
            "actual_position",
            "received_position",
            "position_observed",
            "position_recue",
            "observed_index",
            "actual_index",
        )

    expected = _get(raw, "expected", "attendu", "expected_value", default=_UNSET)
    if expected is _UNSET:
        expected = _detail_value(details, "expected", "attendu", "expected_value")
    observed = _get(
        raw, "observed", "actual", "received", "observe", "recu", "observed_value", default=_UNSET
    )
    if observed is _UNSET:
        observed = _detail_value(
            details, "observed", "actual", "received", "observe", "recu", "observed_value"
        )

    impact = _integer(_get(raw, "impact", "count", "affected_count", default=1))
    if impact is None:
        impact = 1
    return _Issue(
        category=category,
        code=code,
        message=message,
        sheet=_get(raw, "sheet", "sheet_name", "worksheet", "feuille", "onglet", default=None),
        element=_get(raw, "element", "item", "name", "header", "label", default=None),
        expected_position=expected_position,
        observed_position=observed_position,
        expected=expected,
        observed=observed,
        severity=_token(_get(raw, "severity", "niveau", "level", default="warning")) or "warning",
        impact=impact,
        details=details,
    )


def _collect_issues(raw: Any) -> list[_Issue]:
    combined = _get(
        raw,
        "anomalies",
        "issues",
        "findings",
        "differences",
        "structural_anomalies",
        default=_UNSET,
    )
    issues: list[_Issue] = []
    combined_items: list[Any] = []
    if combined is not _UNSET:
        combined_items = _list(combined)
        issues.extend(_normalise_issue(item) for item in combined_items)

    # Certains appelants exposent uniquement une liste par catégorie. On ne les
    # concatène que si aucune liste générale n'existe, afin d'éviter les doublons.
    if not combined_items:
        aliases = {
            "feuilles": ("sheet_anomalies", "sheet_issues", "feuille_anomalies"),
            "colonnes": ("column_anomalies", "column_issues", "colonne_anomalies"),
            "lignes": ("row_anomalies", "row_issues", "ligne_anomalies"),
            "fichier": ("file_anomalies", "file_issues", "fichier_anomalies"),
        }
        for category, names in aliases.items():
            value = _get(raw, *names, default=_UNSET)
            if value is not _UNSET:
                issues.extend(_normalise_issue(item, category) for item in _list(value))
    return issues


def _error_text(value: Any) -> str:
    if isinstance(value, Mapping) or not isinstance(value, (str, bytes, bytearray, BaseException)):
        message = _get(value, "message", "error", "description", default=_UNSET)
        if message is not _UNSET and message is not value:
            code = _text(_get(value, "code", "type", default=""))
            text = _text(message, "Erreur non détaillée")
            return f"{code} — {text}" if code else text
    return _text(value, "Erreur non détaillée")


def _messages(raw: Any, *names: str) -> list[str]:
    messages: list[str] = []
    seen: set[str] = set()
    for name in names:
        found, value = _lookup(raw, name)
        if not found:
            continue
        for item in _list(value):
            text = _error_text(item).strip()
            if text and text not in seen:
                seen.add(text)
                messages.append(text)
    return messages


def _path_country(value: Any) -> str:
    raw = _text(value).strip()
    if not raw:
        return ""
    try:
        # ``Path`` sous Windows gère également les chemins du projet utilisés ici.
        return Path(raw).stem
    except (TypeError, ValueError, OSError):
        return ""


def _provided_count(raw: Any, category: str) -> Optional[int]:
    aliases = {
        "feuilles": ("feuilles", "sheets", "sheet_count", "sheet_anomalies_count"),
        "colonnes": ("colonnes", "columns", "column_count", "column_anomalies_count"),
        "lignes": ("lignes", "rows", "row_count", "row_anomalies_count"),
        "fichier": ("fichier", "file", "files", "file_count", "file_anomalies_count"),
        "autres": ("autres", "other", "others", "other_count"),
    }[category]
    containers = (raw, _get(raw, "counts", default=None), _get(raw, "summary", "resume", default=None))
    for container in containers:
        if container is None:
            continue
        number = _integer(_get(container, *aliases, default=None))
        if number is not None:
            return number
    return None


def _explicit_total(raw: Any) -> Optional[int]:
    for container in (raw, _get(raw, "counts", default=None), _get(raw, "summary", "resume", default=None)):
        if container is None:
            continue
        value = _integer(
            _get(
                container,
                "total_anomalies",
                "anomaly_count",
                "total_issues",
                "total",
                default=None,
            )
        )
        if value is not None:
            return value
    return None


def _is_empty_path(value: Any) -> bool:
    return value is None or not _text(value).strip()


def _truthy_flag(raw: Any, *names: str) -> bool:
    value = _get(raw, *names, default=False)
    if isinstance(value, str):
        return _token(value) in {"1", "true", "yes", "oui", "missing", "manquant"}
    return bool(value)


def _status(
    raw: Any,
    total_anomalies: int,
    errors: Sequence[str],
    reference_path: Any,
    received_path: Any,
) -> tuple[str, str, str, str]:
    explicit = _token(_get(raw, "status", "state", "statut", default=""))
    missing_reference = _truthy_flag(
        raw, "missing_reference", "reference_missing", "sent_missing", "missing_sent", "sans_reference"
    )
    missing_received = _truthy_flag(
        raw, "missing_received", "received_missing", "returned_missing", "missing_returned"
    )
    has_reference_field = _has(
        raw, "reference_path", "reference_file", "sent_path", "sent_file", "fichier_reference", "fichier_envoye"
    )
    has_received_field = _has(
        raw, "received_path", "received_file", "returned_path", "returned_file", "fichier_recu"
    )
    if has_reference_field and _is_empty_path(reference_path):
        missing_reference = True
    if has_received_field and _is_empty_path(received_path):
        missing_received = True

    without_reference_tokens = {
        "sans_reference",
        "reference_missing",
        "missing_reference",
        "unmatched_received",
        "no_reference",
    }
    missing_tokens = {
        "fichier_manquant",
        "missing",
        "missing_file",
        "file_missing",
        "received_missing",
        "missing_received",
        "manquant",
    }
    error_tokens = {
        "erreur",
        "error",
        "failed",
        "failure",
        "invalid",
        "invalide",
        "unreadable",
        "illisible",
        "analysis_error",
    }
    anomaly_tokens = {
        "anomalies",
        "anomaly",
        "non_conforme",
        "noncompliant",
        "warning",
        "ko",
    }

    # Les états rouges explicites sont autoritaires. C'est notamment important
    # pour distinguer un retour sans référence d'un fichier reçu manquant.
    if explicit in without_reference_tokens:
        reason = "Aucun fichier de référence correspondant n'a été trouvé."
        return "sans_reference", "Sans référence", "danger", reason
    if explicit in missing_tokens:
        if missing_reference and missing_received:
            reason = "Les fichiers de référence et reçu sont manquants."
        elif missing_received:
            reason = "Le fichier reçu est manquant."
        else:
            reason = "Un fichier attendu est manquant."
        return "fichier_manquant", "Fichier manquant", "danger", reason
    if explicit in error_tokens:
        return "erreur", "Erreur", "danger", "L'analyse n'a pas pu être menée complètement."
    if missing_reference and missing_received:
        return (
            "fichier_manquant",
            "Fichier manquant",
            "danger",
            "Les fichiers de référence et reçu sont manquants.",
        )
    if missing_reference:
        return (
            "sans_reference",
            "Sans référence",
            "danger",
            "Aucun fichier de référence correspondant n'a été trouvé.",
        )
    if missing_received:
        return "fichier_manquant", "Fichier manquant", "danger", "Le fichier reçu est manquant."
    if errors:
        return "erreur", "Erreur", "danger", "L'analyse n'a pas pu être menée complètement."
    if explicit in anomaly_tokens or total_anomalies:
        return "anomalies", "Anomalies", "warning", ""
    return "conforme", "Conforme", "success", ""


def _normalise_country(raw: Any, hint: str = "", index: int = 0) -> _Country:
    reference_path = _get(
        raw,
        "reference_path",
        "reference_file",
        "sent_path",
        "sent_file",
        "source_path",
        "original_file",
        "fichier_reference",
        "fichier_envoye",
        default=None,
    )
    received_path = _get(
        raw,
        "received_path",
        "received_file",
        "returned_path",
        "returned_file",
        "target_path",
        "actual_file",
        "fichier_recu",
        default=None,
    )
    country = _text(
        _get(raw, "country", "pays", "country_name", "name", default=""),
        "",
    ).strip()
    if not country:
        country = hint.strip() or _path_country(reference_path) or _path_country(received_path)
    if not country:
        country = f"Pays {index + 1}"
    key = _text(_get(raw, "key", "id", "country_code", default=country), country)

    anomalies = _collect_issues(raw)
    errors = _messages(raw, "errors", "error", "analysis_errors")
    warnings = _messages(raw, "warnings", "warning_messages", "avertissements")

    counts = {category: 0 for category in _CATEGORY_ORDER}
    for issue in anomalies:
        counts[issue.category] += max(0, issue.impact)
    for category in _CATEGORY_ORDER:
        provided = _provided_count(raw, category)
        if provided is not None:
            counts[category] = max(counts[category], provided)

    counted_total = sum(counts.values())
    explicit_total = _explicit_total(raw)
    total_anomalies = max(counted_total, explicit_total or 0)
    if total_anomalies > counted_total:
        counts["autres"] += total_anomalies - counted_total

    status_code, status_label, tone, missing_reason = _status(
        raw, total_anomalies, errors, reference_path, received_path
    )
    metadata = _get(raw, "metadata", "meta", default={})
    if not isinstance(metadata, Mapping):
        metadata = {}

    return _Country(
        key=key,
        country=country,
        reference_path=reference_path,
        received_path=received_path,
        status_code=status_code,
        status_label=status_label,
        tone=tone,
        anomalies=anomalies,
        errors=errors,
        warnings=warnings,
        sheet_order_expected=_list(
            _get(raw, "sheet_order_expected", "expected_sheet_order", default=None)
        ),
        sheet_order_observed=_list(
            _get(raw, "sheet_order_observed", "observed_sheet_order", "received_sheet_order", default=None)
        ),
        metadata=metadata,
        counts=counts,
        total_anomalies=total_anomalies,
        missing_reason=missing_reason,
    )


def _looks_like_result(mapping: Mapping[Any, Any]) -> bool:
    keys = {_token(key) for key in mapping}
    return bool(keys & _RESULT_FIELDS)


def _input_items(results: Any) -> list[tuple[str, Any]]:
    if results is None:
        return []

    nested = _get(results, "results", "country_results", default=_UNSET)
    if nested is not _UNSET and nested is not results:
        results = nested

    if isinstance(results, Mapping):
        if _looks_like_result(results):
            return [("", results)]
        return [(_text(key), value) for key, value in results.items()]
    if isinstance(results, (str, bytes, bytearray, Path)):
        return [("", results)]
    try:
        return [("", item) for item in results]
    except TypeError:
        return [("", results)]


def _normalise_results(results: Any) -> list[_Country]:
    normalised: list[_Country] = []
    for index, (hint, raw) in enumerate(_input_items(results)):
        try:
            normalised.append(_normalise_country(raw, hint, index))
        except Exception as exc:  # Un résultat malformé ne doit pas bloquer les autres.
            country = hint.strip() or f"Pays {index + 1}"
            normalised.append(
                _Country(
                    key=country,
                    country=country,
                    reference_path=None,
                    received_path=None,
                    status_code="erreur",
                    status_label="Erreur",
                    tone="danger",
                    anomalies=[],
                    errors=[f"Résultat illisible : {_text(exc, 'erreur inconnue')}"],
                    warnings=[],
                    sheet_order_expected=[],
                    sheet_order_observed=[],
                    metadata={},
                    counts={category: 0 for category in _CATEGORY_ORDER},
                    total_anomalies=0,
                    missing_reason="Le résultat d'analyse n'a pas pu être interprété.",
                )
            )
    return normalised


def safe_slug(
    value: Any,
    *,
    used: Optional[MutableSet[str]] = None,
    salt: str = "",
    max_length: int = 64,
) -> str:
    """Retourne un slug ASCII sûr et, avec ``used``, sans collision.

    Le premier nom conserve un slug lisible. En cas de collision (accents,
    casse, doublon, ou ``index`` réservé), un suffixe SHA-256 court et
    déterministe est ajouté. ``used`` est mis à jour sur place.
    """

    raw = _text(value, "pays").strip() or "pays"
    normalised = unicodedata.normalize("NFKD", raw)
    ascii_value = normalised.encode("ascii", "ignore").decode("ascii").casefold()
    base = re.sub(r"[^a-z0-9]+", "-", ascii_value).strip("-.") or "pays"
    max_length = max(16, int(max_length))
    base = base[:max_length].rstrip("-.") or "pays"

    # Noms de périphériques Windows réservés, même suivis d'une extension.
    reserved = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))}
    if base in reserved:
        base = f"pays-{base}"

    if used is None:
        return base

    occupied = {item.casefold() for item in used}
    candidate = base
    attempt = 0
    while candidate.casefold() in occupied:
        material = f"{raw}\0{salt}\0{attempt}".encode("utf-8", "surrogatepass")
        suffix = sha256(material).hexdigest()[:10]
        room = max(1, max_length - len(suffix) - 1)
        candidate = f"{base[:room].rstrip('-')}-{suffix}"
        attempt += 1
    used.add(candidate)
    return candidate


slugify_country = safe_slug


def _assign_filenames(countries: Sequence[_Country]) -> None:
    used: set[str] = {"index"}
    for index, country in enumerate(countries):
        salt = f"{country.key}\0{index}\0{_text(country.reference_path)}\0{_text(country.received_path)}"
        country.filename = f"{safe_slug(country.country, used=used, salt=salt)}.html"


_STYLES = r"""
:root {
  --ink: #142238; --muted: #657289; --line: #dce3ec; --surface: #ffffff;
  --canvas: #f3f6fa; --brand: #143b66; --brand-2: #21659d;
  --success: #147a4b; --success-bg: #e9f7f0; --success-line: #a9ddc4;
  --warning: #a05208; --warning-bg: #fff4df; --warning-line: #f2c97d;
  --danger: #b42332; --danger-bg: #ffedef; --danger-line: #f1b2ba;
  --shadow: 0 10px 28px rgba(26, 49, 79, .08); --radius: 14px;
}
* { box-sizing: border-box; }
html { color-scheme: light; scroll-behavior: smooth; }
body { margin: 0; color: var(--ink); background: var(--canvas); font: 15px/1.55 Inter, ui-sans-serif, system-ui, -apple-system, "Segoe UI", sans-serif; }
a { color: var(--brand-2); text-decoration-thickness: 1px; text-underline-offset: 3px; }
a:hover { color: var(--brand); }
.topbar { color: #fff; background: linear-gradient(120deg, #0d2d50, #185b8e); box-shadow: 0 3px 18px rgba(14, 46, 77, .22); }
.topbar-inner, main { width: min(1240px, calc(100% - 36px)); margin: 0 auto; }
.topbar-inner { padding: 30px 0 34px; }
.eyebrow { margin: 0 0 5px; color: #b9d9f2; font-size: 12px; font-weight: 800; letter-spacing: .1em; text-transform: uppercase; }
h1 { margin: 0; font-size: clamp(25px, 4vw, 40px); line-height: 1.17; letter-spacing: -.025em; }
.subtitle { max-width: 780px; margin: 9px 0 0; color: #dcecf8; }
.run-meta { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 15px; }
.run-meta span { padding: 4px 9px; border: 1px solid rgba(255,255,255,.23); border-radius: 999px; color: #eaf5fc; background: rgba(255,255,255,.08); font-size: 12px; }
main { padding: 28px 0 48px; }
.kpi-grid { display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 13px; margin-bottom: 24px; }
.kpi { position: relative; min-height: 126px; padding: 19px; overflow: hidden; border: 1px solid var(--line); border-radius: var(--radius); background: var(--surface); box-shadow: var(--shadow); }
.kpi::after { position: absolute; right: -25px; bottom: -36px; width: 92px; height: 92px; border-radius: 50%; background: #eaf1f8; content: ""; }
.kpi.success::after { background: var(--success-bg); } .kpi.warning::after { background: var(--warning-bg); } .kpi.danger::after { background: var(--danger-bg); }
.kpi-label { display: block; min-height: 38px; color: var(--muted); font-size: 12px; font-weight: 750; letter-spacing: .035em; text-transform: uppercase; }
.kpi-value { position: relative; z-index: 1; display: block; margin-top: 4px; font-size: 32px; font-weight: 820; line-height: 1; }
.panel { margin-bottom: 20px; overflow: hidden; border: 1px solid var(--line); border-radius: var(--radius); background: var(--surface); box-shadow: var(--shadow); }
.panel-header { display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 19px 21px; border-bottom: 1px solid var(--line); }
.panel-header h2, .section-title { margin: 0; font-size: 19px; }
.panel-body { padding: 21px; }
.toolbar { display: flex; flex-wrap: wrap; align-items: center; gap: 9px; }
.search { min-width: 245px; padding: 9px 12px; border: 1px solid #bdc9d8; border-radius: 9px; color: var(--ink); background: #fff; font: inherit; }
.search:focus { outline: 3px solid rgba(33,101,157,.18); border-color: var(--brand-2); }
.filter { padding: 8px 11px; border: 1px solid var(--line); border-radius: 9px; color: var(--muted); background: #fff; font: 700 12px/1.2 inherit; cursor: pointer; }
.filter[aria-pressed="true"] { color: #fff; border-color: var(--brand); background: var(--brand); }
.table-wrap { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; }
th { padding: 12px 14px; color: #58677c; background: #f6f8fb; font-size: 11px; font-weight: 800; letter-spacing: .045em; text-align: left; text-transform: uppercase; white-space: nowrap; }
th.num, td.num { text-align: right; }
td { padding: 14px; border-top: 1px solid var(--line); vertical-align: middle; }
tbody tr[data-href] { cursor: pointer; transition: background .15s ease; }
tbody tr[data-href]:hover, tbody tr[data-href]:focus { outline: none; background: #f2f7fb; }
.country-link { color: var(--ink); font-weight: 750; }
.status { display: inline-flex; align-items: center; gap: 7px; padding: 5px 9px; border: 1px solid; border-radius: 999px; font-size: 12px; font-weight: 800; white-space: nowrap; }
.status::before { width: 7px; height: 7px; border-radius: 50%; background: currentColor; content: ""; }
.status.success { color: var(--success); border-color: var(--success-line); background: var(--success-bg); }
.status.warning { color: var(--warning); border-color: var(--warning-line); background: var(--warning-bg); }
.status.danger { color: var(--danger); border-color: var(--danger-line); background: var(--danger-bg); }
.detail-link { font-weight: 750; white-space: nowrap; }
.empty { padding: 44px 22px; color: var(--muted); text-align: center; }
.empty strong { display: block; margin-bottom: 4px; color: var(--ink); font-size: 17px; }
.breadcrumb { display: inline-flex; margin-bottom: 12px; color: #dcecf8; font-weight: 700; }
.summary-strip { display: flex; align-items: center; justify-content: space-between; gap: 18px; margin-bottom: 19px; padding: 18px 20px; border: 1px solid var(--line); border-left: 5px solid var(--brand); border-radius: var(--radius); background: #fff; box-shadow: var(--shadow); }
.summary-strip.success { border-left-color: var(--success); } .summary-strip.warning { border-left-color: var(--warning); } .summary-strip.danger { border-left-color: var(--danger); }
.summary-strip p { margin: 4px 0 0; color: var(--muted); }
.file-grid, .category-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 13px; }
.file-card { padding: 16px; border: 1px solid var(--line); border-radius: 11px; background: #f9fbfd; }
.file-label { display: block; margin-bottom: 7px; color: var(--muted); font-size: 11px; font-weight: 800; letter-spacing: .055em; text-transform: uppercase; }
.file-path { display: block; overflow-wrap: anywhere; color: #243852; font: 13px/1.55 ui-monospace, SFMono-Regular, Consolas, monospace; }
.category-grid { grid-template-columns: repeat(5, minmax(0, 1fr)); }
.category-card { padding: 14px; border: 1px solid var(--line); border-radius: 11px; background: #fff; }
.category-card span { display: block; color: var(--muted); font-size: 12px; }
.category-card strong { display: block; margin-top: 3px; font-size: 25px; }
.alert { margin-bottom: 14px; padding: 14px 16px; border: 1px solid; border-radius: 11px; }
.alert h2 { margin: 0 0 7px; font-size: 16px; }
.alert ul { margin: 0; padding-left: 20px; }
.alert.danger { color: #781c27; border-color: var(--danger-line); background: var(--danger-bg); }
.alert.warning { color: #77400c; border-color: var(--warning-line); background: var(--warning-bg); }
.order-block { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
.order-column { padding: 15px; border: 1px solid var(--line); border-radius: 10px; background: #f9fbfd; }
.order-column h3 { margin: 0 0 8px; font-size: 13px; }
.order-flow { margin: 0; color: #34475f; overflow-wrap: anywhere; }
.issue-group { margin-top: 19px; }
.issue-group-header { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 10px; }
.issue-group-header h2 { margin: 0; font-size: 18px; }
.count-badge { min-width: 29px; padding: 3px 9px; border-radius: 999px; color: var(--brand); background: #e7f0f8; font-size: 12px; font-weight: 800; text-align: center; }
.issue-list { display: grid; gap: 10px; }
.issue { padding: 16px 17px; border: 1px solid var(--line); border-left: 4px solid var(--warning); border-radius: 11px; background: #fff; box-shadow: 0 4px 14px rgba(26,49,79,.045); }
.issue.severity-error, .issue.severity-critical { border-left-color: var(--danger); }
.issue-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; }
.issue h3 { margin: 0; font-size: 15px; overflow-wrap: anywhere; }
.code { display: inline-block; margin-bottom: 3px; color: var(--muted); font: 700 11px/1.3 ui-monospace, SFMono-Regular, Consolas, monospace; letter-spacing: .025em; }
.impact { color: var(--muted); font-size: 12px; white-space: nowrap; }
.issue-message { margin: 9px 0 0; color: #3c4d63; }
.facts { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px 16px; margin: 13px 0 0; }
.fact { min-width: 0; padding-top: 8px; border-top: 1px dashed var(--line); }
.fact dt { color: var(--muted); font-size: 11px; font-weight: 800; letter-spacing: .035em; text-transform: uppercase; }
.fact dd { margin: 2px 0 0; overflow-wrap: anywhere; }
.technical { margin-top: 11px; color: var(--muted); }
.technical summary { cursor: pointer; font-size: 12px; font-weight: 750; }
.technical dl { display: grid; grid-template-columns: minmax(120px, .35fr) 1fr; gap: 6px 12px; margin: 10px 0 0; font-size: 12px; }
.technical dt { font-weight: 750; overflow-wrap: anywhere; } .technical dd { margin: 0; overflow-wrap: anywhere; }
.footer { margin-top: 28px; color: var(--muted); font-size: 12px; text-align: center; }
.sr-only { position: absolute; width: 1px; height: 1px; padding: 0; overflow: hidden; clip: rect(0,0,0,0); white-space: nowrap; border: 0; }
@media (max-width: 1050px) { .kpi-grid { grid-template-columns: repeat(3, 1fr); } .category-grid { grid-template-columns: repeat(3, 1fr); } }
@media (max-width: 700px) {
  .topbar-inner, main { width: min(100% - 22px, 1240px); } .topbar-inner { padding: 23px 0 26px; }
  main { padding-top: 18px; } .kpi-grid { grid-template-columns: repeat(2, 1fr); gap: 9px; }
  .kpi { min-height: 108px; padding: 15px; } .kpi-value { font-size: 27px; }
  .panel-header, .summary-strip { align-items: flex-start; flex-direction: column; }
  .toolbar, .search { width: 100%; } .file-grid, .order-block { grid-template-columns: 1fr; }
  .category-grid { grid-template-columns: repeat(2, 1fr); } .facts { grid-template-columns: 1fr; }
  th, td { padding: 11px 10px; } .hide-mobile { display: none; }
}
@media print {
  body { background: #fff; } .topbar { background: #173e64 !important; print-color-adjust: exact; }
  .toolbar { display: none; } .panel, .kpi, .summary-strip { box-shadow: none; }
  .issue { break-inside: avoid; } main { width: 100%; }
}
"""


_SCRIPT = r"""
document.addEventListener('DOMContentLoaded', function () {
  const table = document.getElementById('country-summary');
  if (table) {
    const rows = Array.from(table.querySelectorAll('tbody tr[data-status]'));
    const search = document.getElementById('country-search');
    const filters = Array.from(document.querySelectorAll('[data-filter]'));
    let active = 'all';
    const apply = function () {
      const query = (search ? search.value : '').trim().toLocaleLowerCase('fr');
      rows.forEach(function (row) {
        const status = row.dataset.status || '';
        const statusMatch = active === 'all' || status === active ||
          (active === 'danger' && (status === 'erreur' || status === 'fichier_manquant' || status === 'sans_reference'));
        const textMatch = !query || row.textContent.toLocaleLowerCase('fr').includes(query);
        row.hidden = !(statusMatch && textMatch);
      });
    };
    if (search) search.addEventListener('input', apply);
    filters.forEach(function (button) {
      button.addEventListener('click', function () {
        active = button.dataset.filter || 'all';
        filters.forEach(function (candidate) { candidate.setAttribute('aria-pressed', String(candidate === button)); });
        apply();
      });
    });
    rows.forEach(function (row) {
      row.tabIndex = 0;
      const go = function () { if (row.dataset.href) window.location.href = row.dataset.href; };
      row.addEventListener('click', function (event) { if (!event.target.closest('a,button')) go(); });
      row.addEventListener('keydown', function (event) { if (event.key === 'Enter') go(); });
    });
  }
});
"""


def _document(title: str, body: str, script: str = "") -> str:
    script_tag = f"<script>{script}</script>" if script else ""
    return (
        "<!doctype html>\n"
        '<html lang="fr">\n<head>\n'
        '<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        f"<title>{_html(title)}</title>\n"
        f"<style>{_STYLES}</style>\n"
        "</head>\n"
        f"<body>{body}{script_tag}</body>\n</html>\n"
    )


def _metadata_html(run_metadata: Optional[Mapping[str, Any]]) -> str:
    metadata = run_metadata if isinstance(run_metadata, Mapping) else {}
    generated = _get(metadata, "generated_at", "timestamp", "run_date", default=None)
    if generated is None:
        generated = datetime.now().astimezone().strftime("%d/%m/%Y à %H:%M")
    chips = [f"<span>Généré le {_html(generated)}</span>"]
    aliases = (
        ("reference_dir", "Dossier de référence"),
        ("sent_dir", "Dossier envoyé"),
        ("received_dir", "Dossier reçu"),
        ("duration_seconds", "Durée (s)"),
        ("version", "Version"),
    )
    emitted: set[str] = set()
    for key, label in aliases:
        value = _get(metadata, key, default=None)
        if value is not None and _text(value).strip() and key not in emitted:
            chips.append(f"<span>{_html(label)} : {_html(value)}</span>")
            emitted.add(key)
    return '<div class="run-meta">' + "".join(chips) + "</div>"


def _infer_run_metadata(
    results: Any,
    run_metadata: Optional[Mapping[str, Any]],
) -> Optional[Mapping[str, Any]]:
    """Récupère les métadonnées d'un ``RunSummary`` si nécessaire."""

    if run_metadata is not None:
        return run_metadata
    inferred: dict[str, Any] = {}
    for name in (
        "generated_at",
        "timestamp",
        "run_date",
        "sent_dir",
        "reference_dir",
        "received_dir",
        "duration_seconds",
        "version",
    ):
        found, value = _lookup(results, name)
        if found and value is not None:
            inferred[name] = value
    return inferred or None


def _kpi(identifier: str, label: str, value: int, tone: str = "") -> str:
    return (
        f'<article class="kpi {tone}">'
        f'<span class="kpi-label">{_html(label)}</span>'
        f'<strong class="kpi-value" id="{identifier}">{int(value)}</strong>'
        "</article>"
    )


def _count_cell(country: _Country, category: str) -> str:
    if country.status_code in {"fichier_manquant", "sans_reference"}:
        return '<span aria-label="Non applicable">—</span>'
    return str(country.counts[category])


def _index_row(country: _Country) -> str:
    filename = _html(country.filename)
    return (
        f'<tr data-status="{country.status_code}" data-country-key="{_html(country.key)}" data-href="{filename}">'
        f'<td><a class="country-link" href="{filename}">{_html(country.country)}</a></td>'
        f'<td><span class="status {country.tone}">{_html(country.status_label)}</span></td>'
        f'<td class="num">{_count_cell(country, "feuilles")}</td>'
        f'<td class="num">{_count_cell(country, "colonnes")}</td>'
        f'<td class="num">{_count_cell(country, "lignes")}</td>'
        f'<td class="num"><strong>{_count_cell(country, "autres") if country.status_code in {"fichier_manquant", "sans_reference"} else country.total_anomalies}</strong></td>'
        f'<td><a class="detail-link" href="{filename}" aria-label="Voir le rapport de {_html(country.country)}">Voir →</a></td>'
        "</tr>"
    )


def _render_index(countries: Sequence[_Country], title: str, run_metadata: Optional[Mapping[str, Any]]) -> str:
    total = len(countries)
    compliant = sum(country.status_code == "conforme" for country in countries)
    anomalous = sum(country.status_code == "anomalies" for country in countries)
    missing = sum(country.status_code in {"fichier_manquant", "sans_reference"} for country in countries)
    errors = sum(country.status_code == "erreur" for country in countries)
    anomalies = sum(country.total_anomalies for country in countries)

    rows = "".join(_index_row(country) for country in countries)
    if not rows:
        rows = '<tr><td colspan="7"><div class="empty"><strong>Aucun résultat</strong>Aucun fichier n’a été fourni pour cette exécution.</div></td></tr>'

    body = f"""
<header class="topbar"><div class="topbar-inner">
  <p class="eyebrow">POPS Check · Rapport global</p>
  <h1>{_html(title)}</h1>
  <p class="subtitle">Vue synthétique des contrôles de conformité structurelle des classeurs retournés.</p>
  {_metadata_html(run_metadata)}
</div></header>
<main>
  <section class="kpi-grid" aria-label="Indicateurs clés">
    {_kpi("kpi-total", "Fichiers analysés", total)}
    {_kpi("kpi-conformes", "Fichiers conformes", compliant, "success")}
    {_kpi("kpi-anomalies", "Avec anomalies", anomalous, "warning")}
    {_kpi("kpi-manquants", "Fichiers manquants", missing, "danger")}
    {_kpi("kpi-erreurs", "Erreurs d’analyse", errors, "danger")}
    {_kpi("kpi-total-anomalies", "Total anomalies", anomalies, "warning")}
  </section>
  <section class="panel">
    <div class="panel-header">
      <h2>Récapitulatif par pays</h2>
      <div class="toolbar" aria-label="Filtres du tableau">
        <label class="sr-only" for="country-search">Rechercher un pays</label>
        <input class="search" id="country-search" type="search" placeholder="Rechercher un pays…">
        <button class="filter" type="button" data-filter="all" aria-pressed="true">Tous</button>
        <button class="filter" type="button" data-filter="conforme" aria-pressed="false">Conformes</button>
        <button class="filter" type="button" data-filter="anomalies" aria-pressed="false">Anomalies</button>
        <button class="filter" type="button" data-filter="danger" aria-pressed="false">Erreurs / manquants</button>
      </div>
    </div>
    <div class="table-wrap">
      <table id="country-summary">
        <caption class="sr-only">Résultats de conformité par pays</caption>
        <thead><tr>
          <th scope="col">Pays</th><th scope="col">Statut</th>
          <th class="num" scope="col">Feuilles</th><th class="num" scope="col">Colonnes</th>
          <th class="num" scope="col">Lignes</th><th class="num" scope="col">Total anomalies</th>
          <th scope="col">Détail</th>
        </tr></thead>
        <tbody>{rows}</tbody>
      </table>
    </div>
  </section>
  <p class="footer">Rapport autonome généré par POPS Check.</p>
</main>"""
    return _document(title, body, _SCRIPT)


def _fact(label: str, value: Any) -> str:
    if value is None or value is _UNSET or not _text(value).strip():
        return ""
    return f'<div class="fact"><dt>{_html(label)}</dt><dd>{_html(value)}</dd></div>'


def _severity_class(severity: str) -> str:
    return "severity-error" if severity in {"error", "critical", "danger", "high", "fatal"} else "severity-warning"


def _details_html(details: Mapping[str, Any]) -> str:
    if not details:
        return ""
    rows: list[str] = []
    for key, value in details.items():
        rows.append(f"<dt>{_html(key)}</dt><dd>{_html(value)}</dd>")
    return (
        '<details class="technical"><summary>Informations techniques</summary><dl>'
        + "".join(rows)
        + "</dl></details>"
    )


def _issue_html(issue: _Issue, number: int) -> str:
    facts = "".join(
        (
            _fact("Feuille", issue.sheet),
            _fact("Élément", issue.element),
            _fact("Position attendue", issue.expected_position),
            _fact("Position observée", issue.observed_position),
            _fact("Attendu", issue.expected),
            _fact("Observé", issue.observed),
        )
    )
    facts_html = f'<dl class="facts">{facts}</dl>' if facts else ""
    impact = f'<span class="impact">{issue.impact} éléments</span>' if issue.impact > 1 else ""
    return f"""
<article class="issue {_severity_class(issue.severity)}" data-category="{issue.category}" data-anomaly-code="{_html(issue.code)}" data-impact="{issue.impact}">
  <div class="issue-head"><div><span class="code">#{number} · {_html(issue.code)}</span><h3>{_html(issue.message)}</h3></div>{impact}</div>
  {facts_html}
  {_details_html(issue.details)}
</article>"""


def _group_html(country: _Country, category: str) -> str:
    count = country.counts[category]
    issues = [issue for issue in country.anomalies if issue.category == category]
    if not count and not issues:
        return ""
    cards = "".join(_issue_html(issue, index) for index, issue in enumerate(issues, 1))
    represented = sum(issue.impact for issue in issues)
    if count > represented:
        missing_details = count - represented
        cards += (
            '<div class="empty">'
            f"<strong>{missing_details} anomalie{'s' if missing_details > 1 else ''} sans détail individuel</strong>"
            "Le moteur d’analyse a fourni un décompte agrégé pour cette catégorie."
            "</div>"
        )
    return f"""
<section class="issue-group" id="categorie-{category}" data-category="{category}">
  <div class="issue-group-header"><h2>{_CATEGORY_LABELS[category]}</h2><span class="count-badge">{count}</span></div>
  <div class="issue-list">{cards}</div>
</section>"""


def _alert(kind: str, title: str, messages: Sequence[str]) -> str:
    if not messages:
        return ""
    items = "".join(f"<li>{_html(message)}</li>" for message in messages)
    return f'<section class="alert {kind}"><h2>{_html(title)}</h2><ul>{items}</ul></section>'


def _order_html(country: _Country) -> str:
    if not country.sheet_order_expected and not country.sheet_order_observed:
        return ""
    if country.sheet_order_expected == country.sheet_order_observed:
        return ""
    expected = _html(country.sheet_order_expected) if country.sheet_order_expected else "Non communiqué"
    observed = _html(country.sheet_order_observed) if country.sheet_order_observed else "Non communiqué"
    return f"""
<section class="panel" id="sheet-order">
  <div class="panel-header"><h2>Ordre des feuilles</h2></div>
  <div class="panel-body order-block">
    <div class="order-column"><h3>Attendu</h3><p class="order-flow">{expected}</p></div>
    <div class="order-column"><h3>Observé</h3><p class="order-flow">{observed}</p></div>
  </div>
</section>"""


def _render_country(country: _Country, title: str) -> str:
    category_cards = "".join(
        f'<article class="category-card" data-category="{category}"><span>{_CATEGORY_LABELS[category]}</span><strong id="count-{category}">{country.counts[category]}</strong></article>'
        for category in _CATEGORY_ORDER
    )
    groups = "".join(_group_html(country, category) for category in _CATEGORY_ORDER)
    if not groups:
        groups = (
            '<div class="panel"><div class="empty">'
            '<strong>Aucune anomalie structurelle détectée</strong>'
            "Le classeur reçu conserve la structure du fichier de référence."
            "</div></div>"
        )

    status_message = country.missing_reason
    if not status_message:
        if country.status_code == "conforme":
            status_message = "La structure du fichier reçu est conforme à la référence."
        elif country.status_code == "anomalies":
            status_message = "Des modifications structurelles nécessitent une vérification."
        else:
            status_message = "Consultez les informations ci-dessous."

    body = f"""
<header class="topbar"><div class="topbar-inner">
  <a class="breadcrumb" href="index.html">← Retour au rapport global</a>
  <p class="eyebrow">POPS Check · Rapport détaillé</p>
  <h1>{_html(country.country)}</h1>
  <p class="subtitle">{_html(title)}</p>
</div></header>
<main>
  <section class="summary-strip {country.tone}" data-status="{country.status_code}">
    <div><span class="status {country.tone}" id="status-global">{_html(country.status_label)}</span><p>{_html(status_message)}</p></div>
    <div><span class="kpi-label">Total anomalies</span><strong class="kpi-value" id="total-anomalies">{country.total_anomalies}</strong></div>
  </section>
  {_alert("danger", "Erreurs rencontrées", country.errors)}
  {_alert("warning", "Avertissements", country.warnings)}
  <section class="panel" id="source-files">
    <div class="panel-header"><h2>Fichiers comparés</h2></div>
    <div class="panel-body file-grid">
      <article class="file-card"><span class="file-label">Fichier de référence</span><code class="file-path">{_html(country.reference_path)}</code></article>
      <article class="file-card"><span class="file-label">Fichier reçu</span><code class="file-path">{_html(country.received_path)}</code></article>
    </div>
  </section>
  <section class="panel" id="category-summary">
    <div class="panel-header"><h2>Résumé par catégorie</h2></div>
    <div class="panel-body category-grid">{category_cards}</div>
  </section>
  {_order_html(country)}
  <section aria-label="Détail des anomalies">{groups}</section>
  <p class="footer"><a href="index.html">Retour au rapport global</a> · Rapport autonome généré par POPS Check.</p>
</main>"""
    return _document(f"{country.country} — {title}", body)


def render_index_html(
    results: Any,
    *,
    title: str = "Contrôle de conformité POPS",
    run_metadata: Optional[Mapping[str, Any]] = None,
) -> str:
    """Construit le HTML global en mémoire, sans écrire de fichier."""

    run_metadata = _infer_run_metadata(results, run_metadata)
    countries = _normalise_results(results)
    _assign_filenames(countries)
    return _render_index(countries, title, run_metadata)


def render_country_html(
    result: Any,
    *,
    title: str = "Contrôle de conformité POPS",
) -> str:
    """Construit le HTML détaillé d'un pays en mémoire."""

    country = _normalise_country(result)
    return _render_country(country, title)


def _write(path: Path, content: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(content, encoding="utf-8", errors="replace", newline="\n")
    temporary.replace(path)
    return path


def generate_reports(
    results: Any,
    output_dir: Any = "rapports",
    run_metadata: Optional[Mapping[str, Any]] = None,
    *,
    title: str = "Contrôle de conformité POPS",
) -> Path:
    """Génère le rapport global et tous les rapports pays.

    Args:
        results: itérable de résultats, dictionnaire ``pays -> résultat`` ou
            objet contenant ``results``/``country_results``.
        output_dir: dossier cible (``rapports`` par défaut).
        run_metadata: métadonnées optionnelles du traitement (date, dossiers,
            version). Les champs connus sont affichés dans l'en-tête global.
        title: titre affiché dans les rapports.

    Returns:
        Le chemin de ``index.html``.
    """

    run_metadata = _infer_run_metadata(results, run_metadata)
    destination = Path(output_dir)
    countries = _normalise_results(results)
    _assign_filenames(countries)
    destination.mkdir(parents=True, exist_ok=True)

    # Chaque pays est rendu indépendamment : une donnée inhabituelle est
    # normalisée en erreur locale et ne bloque pas les autres rapports.
    for country in countries:
        _write(destination / country.filename, _render_country(country, title))
    return _write(destination / "index.html", _render_index(countries, title, run_metadata))


def generate_index_report(
    results: Any,
    output_path: Any = Path("rapports") / "index.html",
    *,
    run_metadata: Optional[Mapping[str, Any]] = None,
    title: str = "Contrôle de conformité POPS",
) -> Path:
    """Génère uniquement le rapport global."""

    run_metadata = _infer_run_metadata(results, run_metadata)
    countries = _normalise_results(results)
    _assign_filenames(countries)
    path = Path(output_path)
    return _write(path, _render_index(countries, title, run_metadata))


def generate_country_report(
    result: Any,
    output_path: Any = None,
    *,
    title: str = "Contrôle de conformité POPS",
) -> Path:
    """Génère uniquement le rapport détaillé d'un pays."""

    country = _normalise_country(result)
    if output_path is None:
        output_path = Path("rapports") / f"{safe_slug(country.country)}.html"
    return _write(Path(output_path), _render_country(country, title))


class ReportGenerator:
    """Façade orientée objet pratique pour la CLI et les intégrations."""

    def __init__(
        self,
        output_dir: Any = "rapports",
        *,
        title: str = "Contrôle de conformité POPS",
    ) -> None:
        self.output_dir = Path(output_dir)
        self.title = title

    def generate(
        self,
        results: Any,
        *,
        run_metadata: Optional[Mapping[str, Any]] = None,
    ) -> Path:
        return generate_reports(
            results,
            self.output_dir,
            run_metadata=run_metadata,
            title=self.title,
        )

    def generate_index(
        self,
        results: Any,
        *,
        run_metadata: Optional[Mapping[str, Any]] = None,
    ) -> Path:
        return generate_index_report(
            results,
            self.output_dir / "index.html",
            run_metadata=run_metadata,
            title=self.title,
        )

    def generate_country(self, result: Any, filename: Any = None) -> Path:
        if filename is None:
            country = _normalise_country(result)
            filename = f"{safe_slug(country.country)}.html"
        return generate_country_report(
            result,
            self.output_dir / Path(filename).name,
            title=self.title,
        )
