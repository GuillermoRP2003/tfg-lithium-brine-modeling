from __future__ import annotations

import pandas as pd

import mainret as core


# Refinamiento de PC1:
# - las 10 primeras subetapas representan la antigua primera decima parte
#   de PC1, dividida en 10 extracciones de 1/100 cada una;
# - las 9 subetapas restantes representan las antiguas fracciones 2-10
#   de PC1, de 1/10 cada una.
PC1_REFINED_FRACTIONS = [0.01] * 10 + [0.10] * 9


def build_pc1_refined_stages(stages: pd.DataFrame) -> pd.DataFrame:
    total_fraction = sum(PC1_REFINED_FRACTIONS)
    if abs(total_fraction - 1.0) > 1e-12:
        raise ValueError(f"Las fracciones de PC1 no suman 1.0: {total_fraction}")

    pc1_mask = stages["stage_id"].astype(str).str.upper().str.strip().eq("PC1")
    if not pc1_mask.any():
        raise ValueError("No se encontro la etapa PC1 en stages_base.txt.")

    pc1 = stages.loc[pc1_mask].iloc[0].copy()
    split_rows = []

    for idx, fraction in enumerate(PC1_REFINED_FRACTIONS, start=1):
        row = pc1.copy()
        row["stage_id"] = f"PC1_R{idx:02d}"
        row["split_parent_stage_id"] = "PC1"
        row["split_part"] = idx
        row["split_parts"] = len(PC1_REFINED_FRACTIONS)
        row["split_fraction"] = fraction

        if "surface_m2" in row.index and pd.notna(row["surface_m2"]):
            row["surface_m2"] = float(pc1["surface_m2"]) * fraction
        if "evap_L_s" in row.index and pd.notna(row["evap_L_s"]):
            row["evap_L_s"] = float(pc1["evap_L_s"]) * fraction

        split_rows.append(row)

    return pd.DataFrame(split_rows).reset_index(drop=True)


def run_all_scenarios_pc1_refined() -> None:
    core.validate_paths()

    feed = core.load_feed()
    stages = build_pc1_refined_stages(core.load_stages())
    months_control = core.load_months_control()
    scenarios_control = core.load_scenarios_control()

    core.validate_feed(feed)
    core.validate_stages(stages)
    core.validate_months_control(months_control)
    core.validate_scenarios_control(scenarios_control)

    scenarios_generated = core.expand_scenarios_control(scenarios_control, months_control)

    core.RUNS.mkdir(parents=True, exist_ok=True)
    core.RESULTS.mkdir(parents=True, exist_ok=True)

    run_stamp = core.build_run_stamp()
    run_stamp_refined = f"{run_stamp}_pc1_refinado"

    batch_runs_dir = core.RUNS / run_stamp_refined
    batch_results_dir = core.RESULTS / run_stamp_refined

    batch_runs_dir.mkdir(parents=True, exist_ok=True)
    batch_results_dir.mkdir(parents=True, exist_ok=True)

    all_clean = []
    all_phases = []

    for _, scenario in scenarios_generated.iterrows():
        feed_mod, stages_mod = core.apply_scenario(feed, stages, months_control, scenario)

        clean_df, phases_df = core.run_case(
            scenario=scenario,
            feed=feed_mod,
            stages=stages_mod,
            batch_runs_dir=batch_runs_dir,
        )

        if not clean_df.empty:
            all_clean.append(clean_df)
        if not phases_df.empty:
            all_phases.append(phases_df)

    all_clean_df = pd.concat(all_clean, ignore_index=True) if all_clean else pd.DataFrame()
    all_phases_df = pd.concat(all_phases, ignore_index=True) if all_phases else pd.DataFrame()
    validation_df = core.build_validation_summary(all_clean_df, scenarios_generated)

    summary_df, evaluation_df, geochemical_df, precipitation_df, scaling_risk_df, metadata = core.build_methodological_outputs(
        all_clean_df,
        all_phases_df,
        scenarios_generated,
        validation_df,
        run_stamp_refined,
    )

    core.export_methodological_outputs(
        summary_df,
        evaluation_df,
        geochemical_df,
        precipitation_df,
        scaling_risk_df,
        metadata,
        validation_df,
        batch_results_dir,
    )

    print(f"Resultados PC1 refinado generados en: {batch_results_dir}")


if __name__ == "__main__":
    run_all_scenarios_pc1_refined()
