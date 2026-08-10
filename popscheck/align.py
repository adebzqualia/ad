from __future__ import annotations

import hashlib
from bisect import bisect_left
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from difflib import SequenceMatcher
import re
from typing import Iterable

from .config import AnalysisConfig, SheetRule
from .workbook import SheetSnapshot


TOKEN_SEPARATOR = "§"
TOKEN_WEIGHTS = {
    "F": 4.0,  # formule normalisée exacte
    "P": 6.0,  # topologie de formule, sans offsets de références
    "L": 10.0,  # libellé stable commun (ancre la plus discriminante)
    "M": 4.5,  # fusion
    "T": 4.5,  # table Excel
    "D": 4.0,  # validation de données
    "X": 2.0,  # dimension de ligne/colonne
    "S": 1.0,  # style canonique (signal faible)
}

_RELATIVE_REFERENCE_RE = re.compile(r"R(?:\[[+-]?\d+\]|\d+)C(?:\[[+-]?\d+\]|\d+)")


@dataclass(slots=True)
class AxisItem:
    index: int
    key: str
    tokens: Counter[str]
    label: str = ""
    information: float = 0.0


@dataclass(slots=True)
class AxisAlignment:
    mapping: dict[int, int] = field(default_factory=dict)  # attendu -> reçu
    scores: dict[tuple[int, int], float] = field(default_factory=dict)
    removed: list[int] = field(default_factory=list)
    added: list[int] = field(default_factory=list)
    moved: list[tuple[int, int, float]] = field(default_factory=list)
    ambiguities: list[str] = field(default_factory=list)

    @property
    def observed_to_expected(self) -> dict[int, int]:
        return {observed: expected for expected, observed in self.mapping.items()}


def _base_token(token: str) -> str:
    return token.rsplit(TOKEN_SEPARATOR, 1)[-1]


def _token_weight(token: str) -> float:
    return TOKEN_WEIGHTS.get(_base_token(token).split("|", 1)[0], 1.0)


def _weighted_size(tokens: Counter[str]) -> float:
    return sum(count * _token_weight(token) for token, count in tokens.items())


def _digest(tokens: Counter[str]) -> str:
    if not tokens:
        return "EMPTY"
    serialized = "\x1f".join(f"{token}\x1e{count}" for token, count in sorted(tokens.items()))
    return hashlib.blake2b(serialized.encode("utf-8"), digest_size=16).hexdigest()


def _contextualize(base: str, orthogonal_index: int | None) -> str:
    if orthogonal_index is None:
        return base
    return f"{orthogonal_index}{TOKEN_SEPARATOR}{base}"


def build_axis_items(
    sheet: SheetSnapshot,
    axis: str,
    extent: int,
    rule: SheetRule,
    common_labels: set[str],
    *,
    orthogonal_mapping: dict[int, int] | None = None,
    observed_side: bool = False,
) -> list[AxisItem]:
    """Construit des signatures sparse; aucune valeur numérique n'est utilisée."""
    counters: list[Counter[str]] = [Counter() for _ in range(extent + 1)]
    labels: list[list[str]] = [[] for _ in range(extent + 1)]

    for (row, column), feature in sheet.cells.items():
        if not rule.cell_is_monitored(row, column):
            continue
        index = row if axis == "row" else column
        orthogonal = column if axis == "row" else row
        if not 1 <= index <= extent:
            continue
        canonical_orthogonal: int | None = None
        if orthogonal_mapping is None:
            canonical_orthogonal = None
        elif observed_side:
            canonical_orthogonal = orthogonal_mapping.get(orthogonal)
            if canonical_orthogonal is None:
                # Un axe orthogonal ajouté ne doit pas faire croire que chaque
                # ligne/colonne a elle-même changé.
                continue
        else:
            canonical_orthogonal = orthogonal

        editable = rule.cell_is_editable(row, column)
        if feature.style:
            counters[index][_contextualize(f"S|{feature.style}", canonical_orthogonal)] += 1
        if feature.formula and not editable:
            counters[index][_contextualize(f"F|{feature.formula}", canonical_orthogonal)] += 1
            topology = _RELATIVE_REFERENCE_RE.sub("REF", feature.formula)
            counters[index][_contextualize(f"P|{topology}", canonical_orthogonal)] += 1
        if feature.label and feature.label in common_labels and not editable:
            counters[index][_contextualize(f"L|{feature.label}", canonical_orthogonal)] += 1
            labels[index].append(feature.label)

    dimensions = sheet.row_dimensions if axis == "row" else sheet.column_dimensions
    for index, dimension in dimensions.items():
        if 1 <= index <= extent:
            counters[index][f"X|{dimension}"] += 1

    for structural_range in sheet.ranges:
        if axis == "row":
            start = max(1, structural_range.min_row)
            end = min(extent, structural_range.max_row)
        else:
            start = max(1, structural_range.min_column)
            end = min(extent, structural_range.max_column)
        for index in range(start, end + 1):
            token = structural_range.axis_token(axis, index)
            if token:
                counters[index][token] += 1

    items: list[AxisItem] = []
    for index in range(1, extent + 1):
        tokens = counters[index]
        item_label = " · ".join(dict.fromkeys(labels[index]))[:160]
        items.append(
            AxisItem(
                index=index,
                key=_digest(tokens),
                tokens=tokens,
                label=item_label,
                information=_weighted_size(tokens),
            )
        )
    return items


def axis_similarity(expected: AxisItem, observed: AxisItem) -> float:
    if expected.key == observed.key:
        return 1.0
    if not expected.tokens and not observed.tokens:
        return 1.0
    if not expected.tokens or not observed.tokens:
        return 0.0
    all_tokens = set(expected.tokens) | set(observed.tokens)
    intersection = 0.0
    union = 0.0
    for token in all_tokens:
        weight = _token_weight(token)
        expected_count = expected.tokens.get(token, 0)
        observed_count = observed.tokens.get(token, 0)
        intersection += min(expected_count, observed_count) * weight
        union += max(expected_count, observed_count) * weight
    return intersection / union if union else 0.0


def _lis_positions(values: list[int]) -> set[int]:
    """Positions d'une LIS déterministe, utilisées pour isoler les vrais moves."""
    if not values:
        return set()
    tails: list[int] = []
    tail_positions: list[int] = []
    predecessors = [-1] * len(values)
    for position, value in enumerate(values):
        insertion = bisect_left(tails, value)
        if insertion == len(tails):
            tails.append(value)
            tail_positions.append(position)
        else:
            tails[insertion] = value
            tail_positions[insertion] = position
        if insertion:
            predecessors[position] = tail_positions[insertion - 1]
    selected: set[int] = set()
    cursor = tail_positions[-1]
    while cursor >= 0:
        selected.add(cursor)
        cursor = predecessors[cursor]
    return selected


def _ranked_candidates(
    expected_indexes: Iterable[int],
    observed_indexes: Iterable[int],
    expected_by_index: dict[int, AxisItem],
    observed_by_index: dict[int, AxisItem],
    threshold: float,
) -> list[tuple[float, int, int]]:
    expected_indexes = list(expected_indexes)
    observed_indexes = list(observed_indexes)
    candidates: list[tuple[float, int, int]] = []
    if len(expected_indexes) * len(observed_indexes) <= 400_000:
        candidate_map = {index: observed_indexes for index in expected_indexes}
    else:
        # Cas fortement divergent : ne jamais construire un produit cartésien
        # de dizaines de milliers de lignes. Les clés exactes et les ancres
        # fortes fournissent un index inversé; un voisinage local sert de repli.
        by_key: dict[str, list[int]] = defaultdict(list)
        by_token: dict[str, list[int]] = defaultdict(list)
        for observed_index in observed_indexes:
            item = observed_by_index[observed_index]
            by_key[item.key].append(observed_index)
            for token in item.tokens:
                if _token_weight(token) >= 4.0:
                    by_token[token].append(observed_index)
        observed_set = set(observed_indexes)
        candidate_map: dict[int, list[int]] = {}
        for expected_index in expected_indexes:
            item = expected_by_index[expected_index]
            pool: set[int] = set()
            exact_group = by_key.get(item.key, ())
            if len(exact_group) <= 250:
                pool.update(exact_group)
            for token in item.tokens:
                if _token_weight(token) >= 4.0:
                    token_group = by_token.get(token, ())
                    if len(token_group) <= 250:
                        pool.update(token_group)
            if not pool:
                pool.update(
                    index
                    for index in range(max(1, expected_index - 4), expected_index + 5)
                    if index in observed_set
                )
            candidate_map[expected_index] = sorted(
                pool, key=lambda index: (abs(index - expected_index), index)
            )[:250]
    for expected_index in expected_indexes:
        for observed_index in candidate_map[expected_index]:
            similarity = axis_similarity(
                expected_by_index[expected_index], observed_by_index[observed_index]
            )
            if similarity >= threshold:
                candidates.append((similarity, expected_index, observed_index))
    candidates.sort(key=lambda item: (-item[0], abs(item[1] - item[2]), item[1], item[2]))
    return candidates


def _unique_high_confidence_pairs(
    expected_indexes: set[int],
    observed_indexes: set[int],
    expected_by_index: dict[int, AxisItem],
    observed_by_index: dict[int, AxisItem],
    analysis: AnalysisConfig,
) -> tuple[list[tuple[int, int, float]], list[str]]:
    candidates = _ranked_candidates(
        expected_indexes,
        observed_indexes,
        expected_by_index,
        observed_by_index,
        analysis.min_axis_similarity,
    )
    by_expected: dict[int, list[float]] = defaultdict(list)
    by_observed: dict[int, list[float]] = defaultdict(list)
    expected_key_counts = Counter(expected_by_index[index].key for index in expected_indexes)
    observed_key_counts = Counter(observed_by_index[index].key for index in observed_indexes)
    expected_label_counts = Counter(
        _base_token(token)
        for index in expected_indexes
        for token in expected_by_index[index].tokens
        if _base_token(token).startswith("L|")
    )
    observed_label_counts = Counter(
        _base_token(token)
        for index in observed_indexes
        for token in observed_by_index[index].tokens
        if _base_token(token).startswith("L|")
    )
    for score, expected_index, observed_index in candidates:
        by_expected[expected_index].append(score)
        by_observed[observed_index].append(score)

    selected: list[tuple[int, int, float]] = []
    used_expected: set[int] = set()
    used_observed: set[int] = set()
    ambiguous: list[str] = []
    for score, expected_index, observed_index in candidates:
        if expected_index in used_expected or observed_index in used_observed:
            continue
        expected_scores = by_expected[expected_index]
        observed_scores = by_observed[observed_index]
        expected_gap = score - (expected_scores[1] if len(expected_scores) > 1 else 0.0)
        observed_gap = score - (observed_scores[1] if len(observed_scores) > 1 else 0.0)
        exact_unique = (
            expected_by_index[expected_index].key == observed_by_index[observed_index].key
            and expected_key_counts[expected_by_index[expected_index].key] == 1
            and observed_key_counts[observed_by_index[observed_index].key] == 1
        )
        shared_labels = {
            _base_token(token)
            for token in expected_by_index[expected_index].tokens
            if _base_token(token).startswith("L|")
        } & {
            _base_token(token)
            for token in observed_by_index[observed_index].tokens
            if _base_token(token).startswith("L|")
        }
        unique_label = any(
            expected_label_counts[label] == 1 and observed_label_counts[label] == 1
            for label in shared_labels
        )
        if unique_label or exact_unique or (
            score >= analysis.move_min_similarity
            and
            expected_gap >= analysis.ambiguity_margin
            and observed_gap >= analysis.ambiguity_margin
        ):
            selected.append((expected_index, observed_index, score))
            used_expected.add(expected_index)
            used_observed.add(observed_index)
        elif score >= analysis.move_min_similarity:
            ambiguous.append(
                f"Appariement non discriminant autour des positions {expected_index} et {observed_index}."
            )
    return selected, ambiguous


def _monotonic_pairs(
    expected_indexes: list[int],
    observed_indexes: list[int],
    expected_by_index: dict[int, AxisItem],
    observed_by_index: dict[int, AxisItem],
    threshold: float,
) -> list[tuple[int, int, float]]:
    if not expected_indexes or not observed_indexes:
        return []
    # Les blocs de même longueur sont appariés positionnellement, même si une
    # formule ou un style a changé. C'est un choix volontaire anti-faux-positifs.
    if len(expected_indexes) == len(observed_indexes):
        return [
            (
                expected_index,
                observed_index,
                axis_similarity(expected_by_index[expected_index], observed_by_index[observed_index]),
            )
            for expected_index, observed_index in zip(expected_indexes, observed_indexes)
        ]

    # Pour les blocs de tailles différentes, un petit alignement dynamique
    # monotone maximise la similarité sans créer de déplacements artificiels.
    n, m = len(expected_indexes), len(observed_indexes)
    if n * m > 300_000:
        pair_count = min(n, m)
        return [
            (
                expected_indexes[position],
                observed_indexes[position],
                axis_similarity(
                    expected_by_index[expected_indexes[position]],
                    observed_by_index[observed_indexes[position]],
                ),
            )
            for position in range(pair_count)
        ]
    scores = [[0.0] * (m + 1) for _ in range(n + 1)]
    decisions = [[0] * (m + 1) for _ in range(n + 1)]  # 1 match, 2 skip expected, 3 skip observed
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            similarity = axis_similarity(
                expected_by_index[expected_indexes[i - 1]],
                observed_by_index[observed_indexes[j - 1]],
            )
            match_gain = max(0.0, similarity - threshold + 0.05)
            choices = (
                (scores[i - 1][j - 1] + match_gain, 1),
                (scores[i - 1][j], 2),
                (scores[i][j - 1], 3),
            )
            best_score, decision = max(choices, key=lambda item: (item[0], -item[1]))
            scores[i][j] = best_score
            decisions[i][j] = decision
    pairs: list[tuple[int, int, float]] = []
    i, j = n, m
    while i and j:
        decision = decisions[i][j]
        if decision == 1:
            expected_index = expected_indexes[i - 1]
            observed_index = observed_indexes[j - 1]
            similarity = axis_similarity(expected_by_index[expected_index], observed_by_index[observed_index])
            if similarity >= threshold:
                pairs.append((expected_index, observed_index, similarity))
            i -= 1
            j -= 1
        elif decision == 2:
            i -= 1
        else:
            j -= 1
    pairs.reverse()
    return pairs


def align_axis(
    expected: list[AxisItem],
    observed: list[AxisItem],
    analysis: AnalysisConfig,
) -> AxisAlignment:
    result = AxisAlignment()
    expected_by_index = {item.index: item for item in expected}
    observed_by_index = {item.index: item for item in observed}
    matcher = SequenceMatcher(
        None,
        [item.key for item in expected],
        [item.key for item in observed],
        # Les très longues suites de lignes formatées à l'identique rendent
        # SequenceMatcher quadratique si chaque répétition reste une ancre.
        autojunk=max(len(expected), len(observed)) > 2_000,
    )
    for block in matcher.get_matching_blocks():
        for offset in range(block.size):
            expected_index = expected[block.a + offset].index
            observed_index = observed[block.b + offset].index
            result.mapping[expected_index] = observed_index
            result.scores[(expected_index, observed_index)] = 1.0

    unmatched_expected = set(expected_by_index) - set(result.mapping)
    unmatched_observed = set(observed_by_index) - set(result.mapping.values())

    high_pairs, ambiguities = _unique_high_confidence_pairs(
        unmatched_expected,
        unmatched_observed,
        expected_by_index,
        observed_by_index,
        analysis,
    )
    result.ambiguities.extend(dict.fromkeys(ambiguities))
    for expected_index, observed_index, score in high_pairs:
        result.mapping[expected_index] = observed_index
        result.scores[(expected_index, observed_index)] = score
        unmatched_expected.discard(expected_index)
        unmatched_observed.discard(observed_index)

    # On traite chaque intervalle restant entre deux ancres déjà appariées. Cela
    # maintient la monotonie et localise correctement les insertions/suppressions.
    all_anchors = sorted(
        (expected_index, observed_index)
        for expected_index, observed_index in result.mapping.items()
    )
    anchor_lis = _lis_positions([observed_index for _, observed_index in all_anchors])
    monotone_anchors = [anchor for position, anchor in enumerate(all_anchors) if position in anchor_lis]
    boundaries = [(0, 0)] + monotone_anchors + [(len(expected) + 1, len(observed) + 1)]
    for (left_expected, left_observed), (right_expected, right_observed) in zip(
        boundaries, boundaries[1:]
    ):
        if right_expected <= left_expected or right_observed <= left_observed:
            continue
        expected_segment = sorted(
            index
            for index in unmatched_expected
            if left_expected < index < right_expected
        )
        observed_segment = sorted(
            index
            for index in unmatched_observed
            if left_observed < index < right_observed
        )
        for expected_index, observed_index, score in _monotonic_pairs(
            expected_segment,
            observed_segment,
            expected_by_index,
            observed_by_index,
            analysis.min_axis_similarity,
        ):
            if expected_index not in unmatched_expected or observed_index not in unmatched_observed:
                continue
            result.mapping[expected_index] = observed_index
            result.scores[(expected_index, observed_index)] = score
            unmatched_expected.remove(expected_index)
            unmatched_observed.remove(observed_index)

    result.removed = sorted(unmatched_expected)
    result.added = sorted(unmatched_observed)

    ordered_pairs = sorted(result.mapping.items())
    lis = _lis_positions([observed_index for _, observed_index in ordered_pairs])
    expected_all_label_counts = Counter(
        _base_token(token)
        for item in expected
        for token in item.tokens
        if _base_token(token).startswith("L|")
    )
    observed_all_label_counts = Counter(
        _base_token(token)
        for item in observed
        for token in item.tokens
        if _base_token(token).startswith("L|")
    )
    for position, (expected_index, observed_index) in enumerate(ordered_pairs):
        if position in lis or expected_index == observed_index:
            continue
        score = result.scores.get((expected_index, observed_index), 0.0)
        expected_info = expected_by_index[expected_index].information
        observed_info = observed_by_index[observed_index].information
        shared_labels = {
            _base_token(token)
            for token in expected_by_index[expected_index].tokens
            if _base_token(token).startswith("L|")
        } & {
            _base_token(token)
            for token in observed_by_index[observed_index].tokens
            if _base_token(token).startswith("L|")
        }
        unique_label = any(
            expected_all_label_counts[label] == 1 and observed_all_label_counts[label] == 1
            for label in shared_labels
        )
        if (
            score >= analysis.move_min_similarity or (unique_label and score >= analysis.min_axis_similarity)
        ) and max(expected_info, observed_info) >= 2.0:
            result.moved.append((expected_index, observed_index, score))
        elif analysis.report_ambiguities:
            result.ambiguities.append(
                f"Ordre potentiellement modifié entre {expected_index} et {observed_index}, "
                "mais les signatures ne permettent pas une identification certaine."
            )
    result.ambiguities = list(dict.fromkeys(result.ambiguities))
    return result
