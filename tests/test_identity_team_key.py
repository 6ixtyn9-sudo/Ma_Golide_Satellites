from ingestion.identity import DEFAULT_TEAM_ALIASES, canonical_team_key


def test_team_key_folds_diacritics_and_punctuation_without_fuzzy_shortening():
    assert canonical_team_key("Budućnost VOLI") == canonical_team_key("Buducnost-VOLI")
    assert canonical_team_key("U-BT Cluj-Napoca") == canonical_team_key("U BT Cluj Napoca")
    # The normalization is spelling-only: it does not strip mascot tokens.
    assert canonical_team_key("UMBC Retrievers") != canonical_team_key("UMBC")


def test_team_key_preserves_team_tokens_and_only_ignores_case_and_separators():
    assert canonical_team_key("Real Madrid") == canonical_team_key("REAL-MADRID")
    assert canonical_team_key("Saint Mary's") != canonical_team_key("Saint Joseph's")


def test_normalized_alias_keys_do_not_collide_across_canonical_teams():
    owners = {}
    for canonical, variants in DEFAULT_TEAM_ALIASES.items():
        for value in (canonical, *variants):
            key = canonical_team_key(value)
            assert key not in owners or owners[key] == canonical
            owners[key] = canonical
