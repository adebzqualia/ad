from __future__ import annotations

import argparse
import json
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Sequence

from . import __version__
from .compare import analyze_directories
from .config import AppConfig, load_config
from .models import CountryResult, Status
from .reporting import generate_reports


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="popscheck",
        description=(
            "Compare la structure, les formules et les dépendances des classeurs Excel POPS, "
            "puis génère des rapports HTML."
        ),
    )
    parser.add_argument(
        "--sent",
        default="data/sent",
        help="dossier des fichiers de référence (défaut : data/sent)",
    )
    parser.add_argument(
        "--received",
        default="data/received",
        help="dossier des fichiers retournés (défaut : data/received)",
    )
    parser.add_argument(
        "--reports",
        default="rapports",
        help="dossier de sortie des rapports (défaut : rapports)",
    )
    parser.add_argument(
        "--config",
        default=None,
        help=(
            "fichier TOML de configuration; sans cette option, popscheck.toml "
            "est utilisé s'il existe dans le dossier courant"
        ),
    )
    parser.add_argument(
        "--fail-on-issues",
        action="store_true",
        help="retourne le code 1 si une anomalie, un manque ou une erreur métier est détecté",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="affiche la trace complète en cas d'erreur fatale",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def _resolve_config_path(argument: str | None) -> Path | None:
    if argument:
        return Path(argument)
    conventional = Path("popscheck.toml")
    return conventional if conventional.is_file() else None


def _summary(results: list[CountryResult]) -> dict[str, int]:
    return {
        "total": len(results),
        "conformes": sum(result.status == Status.CONFORME for result in results),
        "anomalies": sum(result.status == Status.ANOMALIES for result in results),
        "manquants": sum(
            result.status in {Status.FICHIER_MANQUANT, Status.SANS_REFERENCE}
            for result in results
        ),
        "erreurs": sum(result.status == Status.ERREUR for result in results),
        "total_anomalies": sum(result.total_anomalies for result in results),
    }


def run(args: argparse.Namespace) -> int:
    started = time.perf_counter()
    config_path = _resolve_config_path(args.config)
    config: AppConfig = load_config(config_path)
    sent_dir = Path(args.sent).resolve()
    received_dir = Path(args.received).resolve()
    reports_dir = Path(args.reports).resolve()

    print(f"POPS Check {__version__}")
    print(f"Références : {sent_dir}")
    print(f"Reçus      : {received_dir}")
    if config_path:
        print(f"Configuration : {config_path.resolve()}")
    print("Analyse en cours…")

    results = analyze_directories(sent_dir, received_dir, config)
    generated_at = datetime.now().astimezone().isoformat(timespec="seconds")
    duration = round(time.perf_counter() - started, 3)
    run_metadata = {
        "sent_dir": str(sent_dir),
        "received_dir": str(received_dir),
        "generated_at": generated_at,
        "duration_seconds": duration,
        "version": __version__,
        "config_path": str(config_path.resolve()) if config_path else None,
    }
    index_path = generate_reports(results, reports_dir, run_metadata=run_metadata)

    reports_dir.mkdir(parents=True, exist_ok=True)
    json_path = reports_dir / "resultats.json"
    payload = {**run_metadata, "results": [result.to_dict() for result in results]}
    json_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=False),
        encoding="utf-8",
    )

    summary = _summary(results)
    print(
        "Terminé : "
        f"{summary['total']} dossier(s), {summary['conformes']} conforme(s), "
        f"{summary['anomalies']} avec anomalie(s), {summary['manquants']} manquant(s), "
        f"{summary['erreurs']} erreur(s), {summary['total_anomalies']} anomalie(s) au total."
    )
    print(f"Rapport : {Path(index_path).resolve()}")
    if args.fail_on_issues and any(result.status != Status.CONFORME for result in results):
        return 1
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return run(args)
    except (FileNotFoundError, PermissionError, ValueError) as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        if args.verbose:
            traceback.print_exc()
        return 2
    except Exception as exc:  # le lot ne doit jamais produire un traceback brut par défaut
        print(f"Erreur fatale inattendue : {type(exc).__name__}: {exc}", file=sys.stderr)
        if args.verbose:
            traceback.print_exc()
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
