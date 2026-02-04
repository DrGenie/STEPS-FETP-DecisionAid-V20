
/* ===================================================
   STEPS FETP India Decision Aid tool
   Next generation script with working tooltips,
   WTP based benefits, sensitivity, Copilot integration and exports
   =================================================== */

/* ===========================
   Global model coefficients
   =========================== */

const MXL_COEFS = {
  ascProgram: 0.168,
  ascOptOut: -0.601,
  tier: {
    frontline: 0.0,
    intermediate: 0.220,
    advanced: 0.487
  },
  career: {
    certificate: 0.0,
    uniqual: 0.017,
    career_path: -0.122
  },
  mentorship: {
    low: 0.0,
    medium: 0.453,
    high: 0.640
  },
  delivery: {
    blended: 0.0,
    inperson: -0.232,
    online: -1.073
  },
  response: {
    30: 0.0,
    15: 0.546,
    7: 0.610
  },
  costPerThousand: -0.005
};

/* ===========================
   Cost templates (combined)
   =========================== */

const COST_TEMPLATES = {
  frontline: {
    combined: {
      id: "frontline_combined",
      label: "Frontline combined template (all institutions)",
      description:
        "Combined frontline cost structure across all institutions using harmonised components and indirect costs including opportunity cost.",
      oppRate: 1.09,
      components: [
        { id: "staff_core", label: "In country programme staff salaries and benefits", directShare: 0.214 },
        { id: "office_equipment", label: "Office equipment for staff and faculty", directShare: 0.004 },
        { id: "office_software", label: "Office software for staff and faculty", directShare: 0.0004 },
        { id: "rent_utilities", label: "Rent and utilities for staff and faculty", directShare: 0.024 },
        { id: "training_materials", label: "Training materials and printing", directShare: 0.0006 },
        { id: "workshops", label: "Workshops and seminars", directShare: 0.107 },
        { id: "travel_in_country", label: "In country travel for faculty, mentors and trainees", directShare: 0.65 }
      ]
    }
  },
  intermediate: {
    combined: {
      id: "intermediate_combined",
      label: "Intermediate combined template (all institutions)",
      description:
        "Combined intermediate cost structure across all institutions using harmonised components and indirect costs including opportunity cost.",
      oppRate: 0.35,
      components: [
        { id: "staff_core", label: "In country programme staff salaries and benefits", directShare: 0.0924 },
        { id: "staff_other", label: "Other salaries and benefits for consultants and advisors", directShare: 0.0004 },
        { id: "office_equipment", label: "Office equipment for staff and faculty", directShare: 0.0064 },
        { id: "office_software", label: "Office software for staff and faculty", directShare: 0.027 },
        { id: "rent_utilities", label: "Rent and utilities for staff and faculty", directShare: 0.0171 },
        { id: "training_materials", label: "Training materials and printing", directShare: 0.0005 },
        { id: "workshops", label: "Workshops and seminars", directShare: 0.0258 },
        { id: "travel_in_country", label: "In country travel for faculty, mentors and trainees", directShare: 0.57 },
        { id: "travel_international", label: "International travel for faculty, mentors and trainees", directShare: 0.1299 },
        { id: "other_direct", label: "Other direct programme expenses", directShare: 0.1302 }
      ]
    }
  },
  advanced: {
    combined: {
      id: "advanced_combined",
      label: "Advanced combined template (all institutions)",
      description:
        "Combined advanced cost structure across all institutions using harmonised components and indirect costs including opportunity cost.",
      oppRate: 0.30,
      components: [
        { id: "staff_core", label: "In country programme staff salaries and benefits", directShare: 0.165 },
        { id: "office_equipment", label: "Office equipment for staff and faculty", directShare: 0.0139 },
        { id: "office_software", label: "Office software for staff and faculty", directShare: 0.0184 },
        { id: "rent_utilities", label: "Rent and utilities for staff and faculty", directShare: 0.0255 },
        { id: "trainee_allowances", label: "Trainee allowances and scholarships", directShare: 0.0865 },
        { id: "trainee_equipment", label: "Trainee equipment such as laptops and internet", directShare: 0.0035 },
        { id: "trainee_software", label: "Trainee software licences", directShare: 0.0017 },
        { id: "training_materials", label: "Training materials and printing", directShare: 0.0024 },
        { id: "workshops", label: "Workshops and seminars", directShare: 0.0188 },
        { id: "travel_in_country", label: "In country travel for faculty, mentors and trainees", directShare: 0.372 },
        { id: "travel_international", label: "International travel for faculty, mentors and trainees", directShare: 0.288 },
        { id: "other_direct", label: "Other direct programme expenses", directShare: 0.0043 }
      ]
    }
  }
};

let COST_CONFIG = null;

/* ===========================
   Epidemiological settings
   =========================== */

const DEFAULT_EPI_SETTINGS = {
  general: {
    planningHorizonYears: 1,
    inrToUsdRate: 83,
    epiDiscountRate: 0.03
  },
  tiers: {
    frontline: {
      completionRate: 0.9,
      outbreaksPerGraduatePerYear: 0.5,
      valuePerGraduate: 0,
      valuePerOutbreak: 4000000
    },
    intermediate: {
      completionRate: 0.9,
      outbreaksPerGraduatePerYear: 0.5,
      valuePerGraduate: 0,
      valuePerOutbreak: 4000000
    },
    advanced: {
      completionRate: 0.9,
      outbreaksPerGraduatePerYear: 0.5,
      valuePerGraduate: 0,
      valuePerOutbreak: 4000000
    }
  }
};

/* Response time multipliers for outbreak benefits */

const RESPONSE_TIME_MULTIPLIERS = {
  "30": 1.0,
  "15": 1.2,
  "7": 1.5
};

/* Tier duration in months */

const TIER_MONTHS = {
  frontline: 3,
  intermediate: 12,
  advanced: 24
};
// Salary-based opportunity cost (additional module)
// - Defaults are editable via Settings tab.
// - Contact day assumptions are documented in Settings and the technical appendix.
// - Counts per cohort: participants = traineesPerCohort; coordinators = 1; faculty = mentorsPerCohort (from mentorship capacity).
const SALARY_OC_DEFAULTS = {
  facultyMonthly: 200000,
  coordinatorMonthly: 70000,
  participantMonthly: 100000
};

const SALARY_OC_DAYS_PER_MONTH = 30;

// Contact day assumptions for blended delivery (days across the full programme duration).
// Fully in-person uses the full programme duration (months).
// Fully online treats contact days as 0 (participants/coordinators), and faculty time as 50% of the non-contact duration.
const SALARY_OC_CONTACT_DAYS_BY_TIER = {
  frontline: 20,
  intermediate: 60,
  advanced: 90
};

function readSalaryInputsFromDom() {
  const facEl = document.getElementById("adv-salary-monthly-faculty");
  const cooEl = document.getElementById("adv-salary-monthly-coordinator");
  const parEl = document.getElementById("adv-salary-monthly-participant");

  const facultyMonthly = facEl ? Number(facEl.value) : SALARY_OC_DEFAULTS.facultyMonthly;
  const coordinatorMonthly = cooEl ? Number(cooEl.value) : SALARY_OC_DEFAULTS.coordinatorMonthly;
  const participantMonthly = parEl ? Number(parEl.value) : SALARY_OC_DEFAULTS.participantMonthly;

  return {
    facultyMonthly: isFinite(facultyMonthly) && facultyMonthly >= 0 ? facultyMonthly : SALARY_OC_DEFAULTS.facultyMonthly,
    coordinatorMonthly: isFinite(coordinatorMonthly) && coordinatorMonthly >= 0 ? coordinatorMonthly : SALARY_OC_DEFAULTS.coordinatorMonthly,
    participantMonthly: isFinite(participantMonthly) && participantMonthly >= 0 ? participantMonthly : SALARY_OC_DEFAULTS.participantMonthly
  };
}


/* ===========================
   STEPS UPGRADE: Cohort capacity constraint
   =========================== */

function parsePositiveInt(value, fallback) {
  const v = Number(value);
  if (!isFinite(v) || v <= 0) return fallback;
  return Math.max(1, Math.floor(v));
}

function getAvailableTrainingSites() {
  const el = document.getElementById("available-training-sites");
  if (!el) return 1;
  const raw = String(el.value || "").trim();
  // Default assumption if empty/invalid: 1 site
  if (!raw) return 1;
  return parsePositiveInt(raw, 1);
}

function computeMaxFeasibleCohorts(tierKey, horizonYears, sites) {
  const t = String(tierKey || "frontline").toLowerCase();
  const durationMonths = (TIER_MONTHS && TIER_MONTHS[t]) ? Number(TIER_MONTHS[t]) : 12;
  const hYears = Number(horizonYears);
  const horizonMonths = isFinite(hYears) ? hYears * 12 : 0;
  const s = parsePositiveInt(sites, 1);
  if (!isFinite(horizonMonths) || horizonMonths <= 0 || !isFinite(durationMonths) || durationMonths <= 0) {
    return 1;
  }
  const maxC = Math.floor((horizonMonths / durationMonths) * s);
  return Math.max(1, maxC);
}

function setCohortCapacityFeedback(maxCohorts) {
  const maxEl = document.getElementById("max-feasible-cohorts");
  if (maxEl) maxEl.textContent = String(maxCohorts);
}

let _cohortCapacityWarnTimeout = null;
function setCohortCapacityWarning(message) {
  const warnEl = document.getElementById("cohort-capacity-warning");
  if (!warnEl) return;
  warnEl.textContent = message || "";
  if (_cohortCapacityWarnTimeout) {
    clearTimeout(_cohortCapacityWarnTimeout);
    _cohortCapacityWarnTimeout = null;
  }
  if (message) {
    _cohortCapacityWarnTimeout = setTimeout(() => {
      warnEl.textContent = "";
      _cohortCapacityWarnTimeout = null;
    }, 4500);
  }
}

function syncPlanningHorizonInputs(years) {
  const v = String(parsePositiveInt(years, 1));
  const configEl = document.getElementById("planning-horizon-config");
  const advEl = document.getElementById("adv-planning-horizon");
  if (configEl && String(configEl.value || "") !== v) configEl.value = v;
  if (advEl && String(advEl.value || "") !== v) advEl.value = v;
  appState.epiSettings.general.planningHorizonYears = Number(v);
}

function enforceCohortCapacityConstraints(opts) {
  const options = opts || {};
  const tierKey = String(options.tierKey || (document.getElementById("program-tier") ? document.getElementById("program-tier").value : "frontline"));
  const source = String(options.source || "");
  const clampIfNeeded = options.clampIfNeeded !== false;

  const cohortsEl = document.getElementById("cohorts");
  const horizonEl = document.getElementById("planning-horizon-config");

  let horizonYears = horizonEl ? Number(horizonEl.value) : Number(appState.epiSettings.general.planningHorizonYears);
  if (!isFinite(horizonYears) || horizonYears <= 0) horizonYears = 1;

  // Keep horizons aligned (config + settings)
  syncPlanningHorizonInputs(horizonYears);

  const sites = getAvailableTrainingSites();
  const maxCohorts = computeMaxFeasibleCohorts(tierKey, horizonYears, sites);

  if (cohortsEl) {
    cohortsEl.setAttribute("max", String(maxCohorts));
  }
  setCohortCapacityFeedback(maxCohorts);

  let clamped = false;
  if (clampIfNeeded && cohortsEl) {
    let current = Number(cohortsEl.value);
    if (!isFinite(current) || current <= 0) current = 1;
    if (current > maxCohorts) {
      current = maxCohorts;
      cohortsEl.value = String(current);
      clamped = true;
    }
  }

  if (clamped) {
    const msg = "Cohorts adjusted to the maximum feasible given tier duration, horizon, and sites.";
    setCohortCapacityWarning(msg);
    if (source && source !== "init") {
      showToast(msg, "warning");
    }
  }

  return { maxCohorts, sites, horizonYears };
}

// Auto-refresh selected outputs when time-horizon or site capacity changes.
// "Apply configuration" remains the primary workflow for most settings,
// but horizon-dependent outputs (e.g., National simulation totals) should update live.
let _autoRecomputeTimer = null;
function scheduleAutoRecomputeScenario(reason) {
  if (_autoRecomputeTimer) {
    clearTimeout(_autoRecomputeTimer);
    _autoRecomputeTimer = null;
  }
  _autoRecomputeTimer = setTimeout(() => {
    try {
      const config = getConfigFromForm();
      const scenario = computeScenario(config);
      ensureScenarioHasExportNotes(scenario);
      appState.currentScenario = scenario;
      refreshAllOutputs(scenario);

      // If the Copilot tab already has generated text, keep it in sync.
      const out = document.getElementById("copilot-prompt-output");
      if (out && out.dataset && out.dataset.hasGenerated === "true") {
        if (typeof window.renderCopilotPromptBundle === "function") {
          window.renderCopilotPromptBundle();
        }
      }
    } catch (e) {
      console.error("Auto recompute failed", reason, e);
    }
  }, 160);
}


/* ===========================
   Copilot interpretation prompt
   =========================== */

const COPILOT_INTERPRETATION_PROMPT = `
You are a senior health economist advising the Ministry of Health and Family Welfare in India, working with World Bank counterparts, on plans to scale up Field Epidemiology Training Programmes. You receive structured outputs from the STEPS FETP India Decision Aid for one configuration that summarises programme design, costs, epidemiological benefits and results from a mixed logit preference study (endorsement and willingness to pay).

Use only the STEPS scenario JSON that follows as your quantitative evidence. Treat all numbers in the JSON as internally consistent. Work in Indian rupees as the main currency and, where helpful, also report values in millions of rupees.

Write a narrative policy brief of roughly three to five A4 pages. Use headings and paragraphs only and do not use bullet points or numbered lists. Suggested section headings are: Background; Scenario description; Preference study evidence and endorsement; Economic costs; Epidemiological effects; Benefit cost results; Distributional and implementation considerations; and Recommendations.

In Background, explain briefly the role of FETP in India, the purpose of the STEPS decision aid and why combining costs, epidemiological benefits and preference study results is useful for ministries of health and finance.

In Scenario description, summarise the configuration reported in the JSON: tier, career incentive, mentorship intensity, delivery mode, outbreak response time, cohort size and number of cohorts, cost per trainee per month and whether opportunity cost of trainee time is included. Use clear language that senior officials can read quickly.

In Preference study evidence and endorsement, interpret the endorsement and opt out rates and the willingness to pay values from the mixed logit preference study. Explain how strong support for this configuration appears to be and what this implies for negotiations between government and partners.

In Economic costs, describe programme cost per cohort, total economic cost per cohort and total economic cost across all cohorts in the planning horizon. Distinguish clearly between financial costs and economic costs that include opportunity cost where this is relevant.

In Epidemiological effects, explain the number of graduates, implied outbreak responses per year and the epidemiological benefit values. Describe how completion rates, response time and values per graduate and per outbreak response combine to give the total indicative epidemiological benefits.

In Benefit cost results, interpret the benefit cost ratios and net present values. State whether the scenario appears favourable on epidemiological benefits alone and on the combination of willingness to pay and epidemiological benefits and what this implies for the strength of the business case.

In Distributional and implementation considerations, discuss any equity, implementation or capacity issues that logically follow from the scenario structure, such as changes in mentorship intensity, delivery mode or tier, without speculating beyond the JSON.

In Recommendations, give a concise narrative judgement on whether this configuration is a strong, moderate or weak candidate for funding. Suggest any simple variations that might improve value for money and note what further analysis or sensitivity checks ministries could request.

Insert one or two compact tables only if they clarify key results, for example a table comparing costs and benefits per cohort and across all cohorts. Refer to each table in the surrounding text so that the brief remains readable without the table.
`;

/* ===========================
   Tooltip content mapping (UI contract)
   =========================== */

const TOOLTIP_LIBRARY = {
  opt_out_alternative: {
    title: "Opt out alternative",
    body:
      "An opt out option where no new FETP training is funded under the scenario being considered. In STEPS this acts as the benchmark of no new FETP investment."
  },
  cost_components: {
    title: "Cost components",
    body:
      "Combined cost components for each tier, covering salary and benefits, travel, training, trainee support and indirect costs including opportunity cost. In STEPS this provides harmonised direct and indirect cost items used in the costing and economic outputs."
  },
  opportunity_cost: {
    title: "Opportunity cost of trainee time",
    body:
      "The value of trainee salary time spent in training instead of normal duties, per trainee per month. In STEPS this is an optional economic cost component that can be included or excluded in the cost calculations."
  },
  preference_model: {
    title: "Preference model",
    body:
      "Mixed logit preference model estimated from the preference study. In STEPS this model is used to predict endorsement and willingness to pay for different FETP configurations."
  },

  result_endorsement: {
    title: "Endorsement rate",
    body:
      "Predicted share of stakeholders who choose the FETP option rather than the opt out alternative. It is calculated from the mixed logit utility indices using a two option logit share: exp(U_program) divided by exp(U_program) plus exp(U_optout). Higher values indicate stronger predicted support for funding the configuration."
  },
  result_optout: {
    title: "Opt out rate",
    body:
      "Predicted share of stakeholders who choose the opt out alternative rather than funding the configuration. It is one hundred minus the endorsement rate in this two option setup. Higher values indicate weaker predicted support for funding the configuration."
  },
  result_wtp_per_trainee: {
    title: "Perceived programme value per trainee per month",
    body:
      "Indicative rupee value per trainee per month implied by the preference model. It is computed by dividing the non cost utility of the configuration by the absolute value of the cost coefficient and scaling to rupees. It summarises how much value stakeholders place on the package for each trainee per month under the model."
  },
  result_wtp_per_cohort: {
    title: "Total willingness to pay per cohort",
    body:
      "Aggregate willingness to pay for one cohort. It is computed as willingness to pay per trainee per month multiplied by the programme duration in months and multiplied by trainees per cohort. It is an indicative value and should be interpreted alongside cost and epidemiological assumptions."
  },
  result_cost_programme: {
    title: "Programme cost per cohort",
    body:
      "Direct financial cost of running one cohort. It is computed as cost per trainee per month multiplied by programme duration in months and multiplied by trainees per cohort. It excludes opportunity cost unless opportunity cost is included in the economic cost concept."
  },
  result_cost_total: {
    title: "Total economic cost per cohort",
    body:
      "Economic cost concept used for benefit cost calculations. It equals programme cost per cohort plus opportunity cost of trainee time when that component is switched on. Higher values increase the cost base that benefits must exceed to generate ratios above one."
  },
  result_npv: {
    title: "Net benefit per cohort",
    body:
      "Difference between discounted outbreak related epidemiological benefit per cohort and total economic cost per cohort under current settings. Positive values indicate benefits exceed costs; negative values indicate costs exceed the outbreak benefit under the current outbreak value, planning horizon and discount rate assumptions."
  },
  result_bcr: {
    title: "Benefit cost ratio per cohort",
    body:
      "Ratio of discounted outbreak related epidemiological benefit per cohort to total economic cost per cohort. Values above one indicate benefits exceed costs under current assumptions; values below one indicate costs exceed the outbreak benefit. The ratio is sensitive to the value per outbreak, response time multiplier, planning horizon and discount rate."
  },
  result_epi_graduates: {
    title: "Graduates (all cohorts)",
    body:
      "Expected number of graduates across all cohorts after applying completion rates and the endorsement share. It is computed from trainees per cohort times completion rate times endorsement share, multiplied by the number of cohorts. It reflects the scale of trained field epidemiologists produced under the configuration."
  },
  result_epi_outbreaks: {
    title: "Outbreak responses per year",
    body:
      "Expected outbreak responses per year at the configured scale, based on graduates, assumed outbreaks handled per graduate per year, and the response time multiplier. Faster response time increases the credited outbreak responses through the multiplier."
  },
  result_epi_benefit: {
    title: "Outbreak benefit per cohort",
    body:
      "Discounted outbreak related benefit per cohort. It is computed as outbreak responses per year per cohort multiplied by value per outbreak, then multiplied by the present value factor implied by the planning horizon and discount rate. It reflects monetary value under epidemiological assumptions rather than observed savings."
  },

  national_total_cost: {
    title: "Total economic cost (national)",
    body:
      "Total economic cost across all cohorts. It equals total economic cost per cohort multiplied by the number of cohorts. It is the main cost input to national scale benefit cost summaries."
  },
  national_total_benefit: {
    title: "Total outbreak benefit (national)",
    body:
      "Total discounted outbreak related benefit across all cohorts. It equals outbreak benefit per cohort multiplied by the number of cohorts. It depends on the value per outbreak, planning horizon, discount rate and the response time multiplier."
  },
  national_net_benefit: {
    title: "Net outbreak related benefit (national)",
    body:
      "Difference between total outbreak benefit across all cohorts and total economic cost across all cohorts. Positive values indicate outbreak benefits exceed economic costs under current assumptions."
  },
  national_bcr: {
    title: "National benefit cost ratio",
    body:
      "Ratio of total outbreak benefit across all cohorts to total economic cost across all cohorts. Values above one indicate outbreak benefits exceed economic costs under current assumptions at national scale."
  },
  national_graduates: {
    title: "Total graduates (national)",
    body:
      "Total expected graduates over the planning horizon at the configured scale. It aggregates across cohorts after applying completion rates and endorsement share."
  },
  national_outbreaks: {
    title: "Outbreak responses per year (national)",
    body:
      "Aggregate outbreak responses per year implied by all graduates across all cohorts, adjusted by the response time multiplier. It is a model based output dependent on outbreaks per graduate per year and the response time multiplier."
  },
  national_total_wtp: {
    title: "Total willingness to pay (national)",
    body:
      "Aggregate willingness to pay across all cohorts implied by the preference model. It equals willingness to pay per cohort multiplied by the number of cohorts. It summarises model implied value and should be interpreted alongside epidemiological and costing outputs."
  }
};

Object.assign(TOOLTIP_LIBRARY, {
  result_endorsement: {
    title: "Endorsement rate",
    body:
      "Predicted share of stakeholders who would endorse funding the configured FETP option rather than choosing the opt out alternative. Calculated from the mixed logit utility for the programme option versus opt out, converted to a probability and expressed as a percent. Higher values indicate stronger stated support for the package under current assumptions."
  },
  result_optout: {
    title: "Opt out rate",
    body:
      "Predicted share of stakeholders who would choose the opt out alternative rather than fund the configured FETP option. It is the complement of the endorsement rate and sums to 100 percent with it. Higher values indicate weaker stated support for the package."
  },
  result_wtp_per_trainee: {
    title: "WTP per trainee per month",
    body:
      "Indicative willingness to pay per trainee per month implied by the preference model. Calculated as non cost utility for the configured option divided by the absolute value of the cost coefficient, scaled to rupees. Higher values indicate higher implied value placed on the package by respondents in the preference study."
  },
  result_wtp_total_cohort: {
    title: "Total WTP per cohort",
    body:
      "Indicative total willingness to pay for one cohort. Calculated as WTP per trainee per month multiplied by programme duration in months for the selected tier and the number of trainees per cohort. This is not a budget, it is an implied value measure from stated preferences."
  },
  result_programme_cost_cohort: {
    title: "Programme cost per cohort",
    body:
      "Direct financial programme cost for one cohort. Calculated as cost per trainee per month multiplied by programme duration in months for the selected tier and the number of trainees per cohort. This excludes opportunity cost unless that component is explicitly added in the total economic cost indicator."
  },
  result_total_cost_cohort: {
    title: "Total economic cost per cohort",
    body:
      "Economic cost for one cohort used in benefit cost and net benefit calculations. Calculated as programme cost per cohort plus the opportunity cost component when the opportunity cost setting is enabled. Opportunity cost is derived from the tier specific combined template rate applied to the programme cost."
  },
  result_net_benefit_cohort: {
    title: "Net outbreak benefit per cohort",
    body:
      "Net benefit comparing outbreak related epidemiological benefits with economic costs for one cohort. Calculated as discounted outbreak benefit per cohort minus total economic cost per cohort. Positive values indicate outbreak related benefits exceed costs under current assumptions."
  },
  result_bcr: {
    title: "Benefit cost ratio per cohort",
    body:
      "Ratio of discounted outbreak related epidemiological benefits to total economic costs for one cohort. Calculated as outbreak benefit per cohort divided by total economic cost per cohort. Values above 1 indicate outbreak benefits exceed costs under current assumptions."
  },
  result_graduates: {
    title: "Graduates",
    body:
      "Expected number of graduates produced across all configured cohorts, adjusted for completion and endorsement. Calculated from trainees per cohort, the tier completion rate, the endorsement share, and the number of cohorts. Higher values indicate a larger trained workforce output under the configured scale up."
  },
  result_outbreak_responses: {
    title: "Outbreak responses per year",
    body:
      "Expected outbreak responses per year at the configured scale, based on graduates and assumptions about outbreaks handled per graduate per year. Calculated using the effective graduates, the outbreaks per graduate per year setting, and the response time multiplier for the selected response time. Higher values increase estimated outbreak related benefits."
  },
  result_epi_benefit: {
    title: "Outbreak related benefit per cohort",
    body:
      "Discounted outbreak related epidemiological benefit for one cohort over the planning horizon. Calculated from expected outbreak responses per year, value per outbreak, and the present value factor implied by the discount rate and planning horizon. This is an indicative monetary benefit driven by settings and assumptions."
  },

  national_total_cost: {
    title: "Total economic cost",
    body:
      "Total economic cost across all configured cohorts over the planning horizon. Calculated as total economic cost per cohort multiplied by the number of cohorts. Interpreted as the aggregate economic resource requirement under the current configuration and cost settings."
  },
  national_total_benefit: {
    title: "Total outbreak related benefit",
    body:
      "Total discounted outbreak related epidemiological benefit aggregated across all configured cohorts. Calculated as outbreak benefit per cohort multiplied by the number of cohorts. This depends on the outbreak value, discount rate, planning horizon, and assumptions about outbreak responses."
  },
  national_net_benefit: {
    title: "National net benefit",
    body:
      "Net outbreak related benefit at national scale. Calculated as total outbreak related benefit across all cohorts minus total economic cost across all cohorts. Positive values indicate outbreak benefits exceed costs at scale under current assumptions."
  },
  national_bcr: {
    title: "National benefit cost ratio",
    body:
      "National scale benefit cost ratio comparing total outbreak related benefits with total economic costs across all cohorts. Calculated as total outbreak related benefit divided by total economic cost. Values above 1 indicate benefits exceed costs at scale under current assumptions."
  },
  national_total_wtp: {
    title: "Total WTP",
    body:
      "Total indicative willingness to pay aggregated across all configured cohorts. Calculated as total WTP per cohort multiplied by the number of cohorts. This is an implied value measure from the preference model, not a financial budget."
  },
  national_graduates: {
    title: "Total graduates",
    body:
      "Total expected graduates across all configured cohorts, adjusted for completion and endorsement. Calculated by aggregating cohort level graduate outputs across cohorts. Higher values reflect larger workforce scale up under the configured programme."
  },
  national_outbreaks_per_year: {
    title: "Outbreak responses per year",
    body:
      "Expected outbreak responses per year at national scale, aggregating the cohort level implied outbreak responses. This depends on graduate outputs, outbreaks per graduate per year, and the response time multiplier. Higher values increase estimated outbreak related benefits."
  }
});


Object.assign(TOOLTIP_LIBRARY, {
  salary_oc_additional: {
    title: "Salary-based opportunity cost (additional)",
    body: "This additional module estimates the value of time that faculty, coordinators and participants spend on training instead of regular duties. It uses average monthly salaries and the duration and contact days set for each programme tier. For fully in person delivery, salary time is counted for the full programme duration in months. For blended delivery, participants and coordinators are counted only for contact days, while faculty are counted for contact days plus half of the remaining non contact duration. Formula (fully in person): OC_role = monthly salary_role x duration_months x number of people. Formula (blended participants and coordinators): OC_role = monthly salary_role x (contact days/30) x number of people. Formula (blended faculty): OC_faculty = monthly salary_faculty x ((contact days/30) + 0.5 x ((total days - contact days)/30)) x number of faculty. This is an indirect economic cost, not a cash payment."
  },
  result_salary_opp_cost: {
    title: "Salary-based opportunity cost (additional)",
    body: "Salary based opportunity cost is calculated in parallel to the existing STEPS opportunity cost method and is added to total economic cost. It reports the value of time diverted from routine work for faculty, coordinators and participants. It varies by tier through programme duration and contact days, and by modality through the blended rules."
  },
  result_cost_total: {
    title: "Total economic cost per cohort",
    body: "Economic cost used in benefit cost calculations. It equals direct financial cost per cohort plus opportunity costs when the opportunity cost switch is on. STEPS includes two opportunity cost components: the existing rate-based method and the additional salary-based method. Both are labelled separately in the results and exports."
  },
  result_total_cost_cohort: {
    title: "Total economic cost per cohort",
    body: "Economic cost used in benefit cost calculations. It equals direct financial cost per cohort plus opportunity costs when the opportunity cost switch is on. STEPS includes two opportunity cost components: the existing rate-based method and the additional salary-based method. Both are labelled separately in the results and exports."
  }
});
/* ===========================
   Global state
   =========================== */

const appState = {
  currency: "INR",
  usdRate: DEFAULT_EPI_SETTINGS.general.inrToUsdRate,
  epiSettings: JSON.parse(JSON.stringify(DEFAULT_EPI_SETTINGS)),
  currentScenario: null,
  savedScenarios: [],
  
  baselineConfig: null,
  baselineScenario: null,
autoScenarioName: true,
  _lastAutoNameSignature: "",
  charts: {
    uptake: null,
    bcr: null,
    epi: null,
    natCostBenefit: null,
    natEpi: null
  },
  tooltip: {
    tooltipEl: null,
    titleEl: null,
    bodyEl: null,
    currentTarget: null,
    hideTimeout: null
  },
  tour: {
    steps: [],
    currentIndex: 0,
    overlayEl: null,
    popoverEl: null
  },
  settings: {
    lastAppliedValues: null
  }
};

/* ===========================
   Utility functions
   =========================== */

function getElByIdCandidates(ids) {
  if (!Array.isArray(ids)) return null;
  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    if (!id) continue;
    const el = document.getElementById(id);
    if (el) return el;
  }
  return null;
}

function formatNumber(x, decimals = 0) {
  if (x === null || x === undefined || isNaN(x)) return "-";
  return x.toLocaleString("en-IN", {
    maximumFractionDigits: decimals,
    minimumFractionDigits: decimals
  });
}

function formatPct(value, decimals = 1) {
  const v = Number(value);
  if (!isFinite(v)) return "";
  const pct = (v > 1.5) ? v : (v * 100);
  const d = (decimals == null) ? 1 : Number(decimals);
  const fixed = isFinite(d) ? d : 1;
  return `${pct.toFixed(fixed)}%`;
}

function safeNumber(value, fallback) {
  const fb = (fallback === undefined) ? 0 : fallback;
  const n = Number(value);
  return (value === null || value === undefined || !isFinite(n)) ? fb : n;
}



function deriveGraduatesByTierForExport(scenario) {
  // Returns graduates by tier for the full planning horizon (all cohorts).
  // Handles both single-tier scenarios and portfolio (baseline / packages) scenarios.
  const out = { total: 0, frontline: 0, intermediate: 0, advanced: 0 };
  if (!scenario) return out;

  const total = safeNumber(scenario.graduatesAllCohorts ?? 0, 0);
  out.total = total;

  const cfg = scenario.config || {};
  const tier = cfg && cfg.tier ? String(cfg.tier) : "";

  // Portfolio baseline or combined packages
  const isPortfolio = !!cfg.isPortfolio || tier === "portfolio" || !!scenario.portfolioBreakdown || !!scenario._portfolioBreakdown;
  if (isPortfolio) {
    const bd = scenario.portfolioBreakdown || scenario._portfolioBreakdown || {};
    const f = safeNumber(bd.frontline?.graduatesAllCohorts ?? 0, 0);
    const i = safeNumber(bd.intermediate?.graduatesAllCohorts ?? 0, 0);
    const a = safeNumber(bd.advanced?.graduatesAllCohorts ?? 0, 0);
    out.frontline = f;
    out.intermediate = i;
    out.advanced = a;
    out.total = f + i + a;
    return out;
  }

  // Single-tier scenarios
  if (tier === "frontline") out.frontline = total;
  if (tier === "intermediate") out.intermediate = total;
  if (tier === "advanced") out.advanced = total;

  return out;
}

function formatExportNotes(notesObj) {
  const en = notesObj && typeof notesObj === "object" ? (String(notesObj.enablers || "").trim()) : "";
  const rk = notesObj && typeof notesObj === "object" ? (String(notesObj.risks || "").trim()) : "";
  if (!en && !rk) return "";
  const parts = [];
  if (en) parts.push(`<h3>Enablers</h3><p>${safeText(en).replace(/\n/g, "<br>")}</p>`);
  if (rk) parts.push(`<h3>Risks and mitigation</h3><p>${safeText(rk).replace(/\n/g, "<br>")}</p>`);
  return parts.join("\n");
}

function refreshScenarioForExport(s) {
  // Recompute scenario outputs from its saved configuration so exports reflect the latest settings/benefit assumptions.
  if (!s || !s.config) return s;
  const keep = { _sid: s._sid, shortlisted: s.shortlisted, pinned: s.pinned, timestamp: s.timestamp, id: s.id, name: s.name };
  const recomputed = computeScenario(s.config);
  if (keep._sid) recomputed._sid = keep._sid;
  if (keep.shortlisted !== undefined) recomputed.shortlisted = keep.shortlisted;
  if (keep.pinned !== undefined) recomputed.pinned = keep.pinned;
  if (keep.name) recomputed.name = keep.name;
  return recomputed;
}


function safeBcr(s) {
  const cost = Number(s.totalEconomicCostAllCohorts ?? s.natTotalCost ?? 0);
  const benefit = Number(s.totalEpiBenefitsAllCohorts ?? s.epiBenefitAllCohorts ?? 0);
  if (cost > 0) return benefit / cost;
  return 0;
}


function formatCurrencyINR(amount, decimals = 0) {
  if (amount === null || amount === undefined || isNaN(amount)) return "-";
  return "INR " + formatNumber(amount, decimals);
}

function formatCurrencyDisplay(amountInINR, decimals = 0) {
  if (amountInINR === null || amountInINR === undefined || isNaN(amountInINR)) return "-";
  if (appState.currency === "USD") {
    const usd = amountInINR / (appState.usdRate || 1);
    return "USD " + formatNumber(usd, decimals);
  }
  return formatCurrencyINR(amountInINR, decimals);
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function presentValueFactor(rate, years) {
  if (years <= 0) return 0;
  if (rate <= 0) return years;
  const r = rate;
  return (1 - Math.pow(1 + r, -years)) / r;
}

function safeText(x) {
  if (x === null || x === undefined) return "";
  return String(x);
}


function cloneScenario(s) {
  // Deep clone for safe storage (avoids reference mutation across saves)
  try {
    const cloned = JSON.parse(JSON.stringify(s));
    if (!cloned) return s;
    if (!cloned._sid) cloned._sid = `sc_${Date.now()}_${Math.floor(Math.random() * 1e6)}`;
    if (cloned.shortlisted === undefined) cloned.shortlisted = false;
    if (cloned.pinned === undefined) cloned.pinned = false;
    return cloned;
  } catch (e) {
    return s;
  }
}





/* ===========================
   Baseline + Planner persistence
   =========================== */

const SAVED_SCENARIOS_STORAGE_KEY = "STEPS_SAVED_SCENARIOS_V2";
const BASELINE_CONFIG_STORAGE_KEY = "STEPS_BASELINE_CONFIG_V1";
const PIN_LIMIT = 3;

let __rid = 0;
function randomId(prefix = "id") {
  __rid += 1;
  return `${prefix}_${Date.now()}_${__rid}`;
}

function signedNumber(n, decimals = 0) {
  const v = Number(n);
  if (!isFinite(v)) return "-";
  const sign = v > 0 ? "+" : v < 0 ? "-" : "";
  return sign + formatNumber(Math.abs(v), decimals);
}

function signedCurrency(n, decimals = 0) {
  const v = Number(n);
  if (!isFinite(v)) return "-";
  const sign = v > 0 ? "+" : v < 0 ? "-" : "";
  return sign + formatCurrencyDisplay(Math.abs(v), decimals);
}

function safeJsonParse(text, fallback) {
  try {
    if (text === null || text === undefined) return fallback;
    const out = JSON.parse(text);
    return out === undefined ? fallback : out;
  } catch (e) {
    return fallback;
  }
}

function normaliseTierKey(tier) {
  const t = String(tier || "").toLowerCase();
  if (t.includes("front")) return "frontline";
  if (t.includes("inter")) return "intermediate";
  if (t.includes("adv")) return "advanced";
  return t;
}

function buildDefaultBaselineConfig() {
  // Baseline (business as usual) defaults used on first load and when the user clicks "Reset baseline to default".
  // Values match the agreed STEPS baseline used for submission-ready reports.
  // We keep both availableTrainingSites and availableSites for backward compatibility.
  return {
    availableTrainingSites: 4,
    availableSites: 4,
    tiers: {
      frontline: {
        cohorts: 36,
        traineesPerCohort: 35,
        costPerTraineePerMonth: 90000,
        completionPct: 94,
        career: "certificate",
        mentorship: "low",
        delivery: "inperson"
      },
      intermediate: {
        cohorts: 11,
        traineesPerCohort: 18,
        costPerTraineePerMonth: 120000,
        completionPct: 80,
        career: "certificate",
        mentorship: "medium",
        delivery: "inperson"
      },
      advanced: {
        cohorts: 14,
        traineesPerCohort: 10,
        costPerTraineePerMonth: 150000,
        completionPct: 80,
        career: "uniqual",
        mentorship: "high",
        delivery: "inperson"
      }
    }
  };
}

function loadBaselineConfig() {
  const raw = safeJsonParse(localStorage.getItem(BASELINE_CONFIG_STORAGE_KEY), null);
  const def = buildDefaultBaselineConfig();
  if (!raw || typeof raw !== "object") return def;
  const out = { ...def, ...raw };

  const sitesFromRaw = out.availableTrainingSites ?? out.availableSites;
  const sites = Math.max(1, Number(sitesFromRaw) || 1);
  out.availableTrainingSites = sites;
  out.availableSites = sites;

  out.tiers = out.tiers || {};
  for (const k of ["frontline","intermediate","advanced"]) {
    const base = def.tiers[k] || {};
    const v = out.tiers[k] || {};
    out.tiers[k] = {
      cohorts: Math.max(0, Math.round(Number(v.cohorts ?? base.cohorts) || 0)),
      traineesPerCohort: Math.max(1, Math.round(Number(v.traineesPerCohort ?? base.traineesPerCohort) || 1)),
      costPerTraineePerMonth: Math.max(0, Number(v.costPerTraineePerMonth ?? base.costPerTraineePerMonth) || 0),
      completionPct: Math.max(0, Math.min(100, Number(v.completionPct ?? base.completionPct) || 0)),
      career: String(v.career ?? base.career ?? "certificate"),
      mentorship: String(v.mentorship ?? base.mentorship ?? "medium"),
      delivery: String(v.delivery ?? base.delivery ?? "blended")
    };
  }

  return out;
}

function persistBaselineConfig(cfg) {
  try {
    localStorage.setItem(BASELINE_CONFIG_STORAGE_KEY, JSON.stringify(cfg));
  } catch (e) {}
}

function persistSavedScenarios() {
  try {
    const payload = appState.savedScenarios.map((s) => ({
      _sid: s._sid || s.id || randomId("sc"),
      name: s.config?.name || s.name || "Scenario",
      createdAt: s.createdAt || Date.now(),
      pinned: !!s.pinned,
      pinnedAt: s.pinnedAt || null,
      shortlisted: !!s.shortlisted,
      config: s.config || {}
    }));
    localStorage.setItem(SAVED_SCENARIOS_STORAGE_KEY, JSON.stringify(payload));
  } catch (e) {}
}

function loadSavedScenarios() {
  const raw = safeJsonParse(localStorage.getItem(SAVED_SCENARIOS_STORAGE_KEY), []);
  if (!Array.isArray(raw)) return [];
  return raw.filter((r) => r && typeof r === "object");
}

function enforcePinLimitDeterministic() {
  const pinned = appState.savedScenarios.filter((s) => !!s.pinned);
  if (pinned.length <= PIN_LIMIT) return;
  // Keep earliest pinned; unpin the rest deterministically
  pinned.sort((a,b) => (a.pinnedAt||0) - (b.pinnedAt||0) || String(a._sid||"").localeCompare(String(b._sid||"")));
  for (const s of pinned.slice(PIN_LIMIT)) {
    s.pinned = false;
    s.pinnedAt = null;
  }
}

function getPinnedScenariosDeterministic() {
  const pinned = appState.savedScenarios.filter((s) => !!s.pinned);
  pinned.sort((a,b) => (a.pinnedAt||0) - (b.pinnedAt||0) || String(a._sid||"").localeCompare(String(b._sid||"")));
  return pinned;
}

function hydrateScenarioRecord(rec) {
  if (!rec || typeof rec !== "object") return null;
  const cfg = rec.config || {};
  let scenario;
  try {
    if (cfg && cfg.isPortfolio && cfg.portfolioTiers) scenario = computePortfolioScenario(cfg);
    else scenario = computeScenario(cfg);
  } catch (e) {
    return null;
  }
  scenario._sid = rec._sid || scenario._sid || randomId("sc");
  scenario.createdAt = rec.createdAt || Date.now();
  scenario.pinned = !!rec.pinned;
  scenario.pinnedAt = rec.pinnedAt || null;
  scenario.shortlisted = !!rec.shortlisted;
  if (scenario.config) {
    scenario.config.name = rec.name || scenario.config.name;
  }
  return scenario;
}

/* ===========================
   Outbreak value presets (sensitivity dropdowns)
   =========================== */

const OUTBREAK_VALUE_PRESETS_MN = [4, 5, 10, 20, 30, 40];

function formatOutbreakPresetLabelMn(mn) {
  return `₹${formatNumber(mn, 0)}m`;
}

function parseSensitivityValueToINR(raw) {
  if (raw === null || raw === undefined) return null;

  if (typeof raw === "number") {
    const n = raw;
    if (!isFinite(n) || n <= 0) return null;
    if (n < 1000) return n * 1e6;
    return n;
  }

  let s = String(raw).trim();
  if (!s) return null;

  const lower = s.toLowerCase().replace(/,/g, "");
  const hasBn = lower.includes("bn") || lower.includes("billion");
  const hasMn = lower.includes("mn") || lower.includes("million");
  const hasCr = lower.includes("crore") || /(^|\s)cr(\s|$)/.test(lower);

  const match = lower.match(/-?\d+(\.\d+)?/);
  if (!match) return null;

  const n = Number(match[0]);
  if (!isFinite(n) || n <= 0) return null;

  if (hasBn) return n * 1e9;
  if (hasMn) return n * 1e6;
  if (hasCr) return n * 1e7;

  if (n < 1000) return n * 1e6;
  return n;
}

function normalisedOutbreakValueKeysFromOption(optionEl) {
  const keys = [];
  if (!optionEl) return keys;

  const rawValue = optionEl.value;
  const rawText = optionEl.textContent || "";

  const inr1 = parseSensitivityValueToINR(rawValue);
  const inr2 = inr1 ? null : parseSensitivityValueToINR(rawText);
  const inr = inr1 || inr2;

  if (rawValue !== null && rawValue !== undefined) keys.push(String(rawValue));

  if (inr && isFinite(inr) && inr > 0) {
    keys.push(String(inr));
    const mn = inr / 1e6;
    if (isFinite(mn)) {
      const mnRounded = Math.round(mn);
      if (Math.abs(mn - mnRounded) < 1e-6) keys.push(String(mnRounded));
      else keys.push(String(mn));
    }
  }

  return keys;
}

function ensureSelectHasOutbreakPresets(selectEl) {
  if (!selectEl) return;

  const existingValues = new Set();
  const hasAnyOptions = selectEl.options && selectEl.options.length > 0;

  Array.from(selectEl.options).forEach((o) => {
    normalisedOutbreakValueKeysFromOption(o).forEach((k) => existingValues.add(String(k)));
  });

  OUTBREAK_VALUE_PRESETS_MN.forEach((mn) => {
    const mnValue = String(mn);
    const inrValue = String(mn * 1e6);

    if (existingValues.has(mnValue) || existingValues.has(inrValue)) return;

    const opt = document.createElement("option");
    opt.value = mnValue;
    opt.textContent = formatOutbreakPresetLabelMn(mn);
    selectEl.appendChild(opt);

    existingValues.add(mnValue);
    existingValues.add(inrValue);
  });

  if (!hasAnyOptions && selectEl.options && selectEl.options.length) {
    const currentInr = appState.epiSettings.tiers.frontline.valuePerOutbreak;
    setSelectToOutbreakValue(selectEl, currentInr);
  }
}

function closestPresetMn(valueInINR) {
  if (!isFinite(valueInINR) || valueInINR <= 0) return OUTBREAK_VALUE_PRESETS_MN[0];
  const mn = valueInINR / 1e6;
  let best = OUTBREAK_VALUE_PRESETS_MN[0];
  let bestDist = Math.abs(best - mn);
  for (let i = 1; i < OUTBREAK_VALUE_PRESETS_MN.length; i++) {
    const v = OUTBREAK_VALUE_PRESETS_MN[i];
    const d = Math.abs(v - mn);
    if (d < bestDist) {
      best = v;
      bestDist = d;
    }
  }
  return best;
}

function setSelectToOutbreakValue(selectEl, valueInINR) {
  if (!selectEl) return;
  if (!isFinite(Number(valueInINR)) || Number(valueInINR) <= 0) return;

  const target = Number(valueInINR);
  const options = Array.from(selectEl.options || []);
  const optionValues = new Set(options.map((o) => String(o.value)));

  const exactInr = String(target);
  if (optionValues.has(exactInr)) {
    selectEl.value = exactInr;
    return;
  }

  let bestOpt = null;
  let bestDist = Infinity;

  options.forEach((opt) => {
    const inr = parseSensitivityValueToINR(opt.value) || parseSensitivityValueToINR(opt.textContent);
    if (!inr || !isFinite(Number(inr))) return;
    const d = Math.abs(Number(inr) - target);
    if (d < bestDist) {
      bestDist = d;
      bestOpt = opt;
    }
  });

  if (bestOpt && isFinite(bestDist)) {
    selectEl.value = bestOpt.value;
    return;
  }

  const nearestMn = closestPresetMn(target);
  const mnCandidate = String(nearestMn);
  const inrCandidate = String(nearestMn * 1e6);

  if (optionValues.has(mnCandidate)) {
    selectEl.value = mnCandidate;
    return;
  }
  if (optionValues.has(inrCandidate)) {
    selectEl.value = inrCandidate;
    return;
  }
}

function syncOutbreakValueDropdownsFromState() {
  const currentInr = appState.epiSettings.tiers.frontline.valuePerOutbreak;

  const sensSelect = getElByIdCandidates(["sensitivityValueSelect", "sensitivity-value-select", "sensitivity-value"]);
  const presetSelect = getElByIdCandidates(["outbreak-value-preset", "outbreakValuePreset", "outbreak-value"]);

  if (sensSelect) {
    ensureSelectHasOutbreakPresets(sensSelect);
    setSelectToOutbreakValue(sensSelect, currentInr);
  }
  if (presetSelect) {
    ensureSelectHasOutbreakPresets(presetSelect);
    setSelectToOutbreakValue(presetSelect, currentInr);
  }
}

function initOutbreakSensitivityDropdowns() {
  const sensSelect = getElByIdCandidates(["sensitivityValueSelect", "sensitivity-value-select", "sensitivity-value"]);
  const presetSelect = getElByIdCandidates(["outbreak-value-preset", "outbreakValuePreset", "outbreak-value"]);

  if (sensSelect) ensureSelectHasOutbreakPresets(sensSelect);
  if (presetSelect) ensureSelectHasOutbreakPresets(presetSelect);

  syncOutbreakValueDropdownsFromState();
}

function enforceResponseTimeFixedTo7Days() {
  const responseEl = document.getElementById("response");
  if (!responseEl || responseEl.tagName.toLowerCase() !== "select") return;

  responseEl.value = "7";

  Array.from(responseEl.options).forEach((opt) => {
    const v = String(opt.value);
    if (v === "15" || v === "30") {
      opt.disabled = true;
      opt.setAttribute("aria-disabled", "true");
    }
    if (v === "7") {
      opt.disabled = false;
      opt.removeAttribute("aria-disabled");
    }
  });

  responseEl.addEventListener("change", () => {
    responseEl.value = "7";
  });
}

/* ===========================
   Toasts (UI contract: #toastContainer)
   =========================== */

function showToast(message, type = "info") {
      const container = document.getElementById("toastContainer");
      if (!container) return;

      const toast = document.createElement("div");
      toast.className = "toast-message";
      toast.setAttribute("role", "status");
      toast.setAttribute("aria-live", "polite");

      const t = type === "success" || type === "warning" || type === "error" ? type : "info";
      toast.dataset.toastType = t;
      toast.classList.add("toast-" + t);

      toast.textContent = message;
      container.appendChild(toast);

      const maxToasts = 4;
      while (container.children.length > maxToasts) {
        container.removeChild(container.firstChild);
      }

      const remove = () => {
        if (toast && toast.parentNode) {
          toast.parentNode.removeChild(toast);
        }
      };

      setTimeout(remove, 3200);
    }

/* ===========================
       Tooltip system (UI contract: #globalTooltip, .tooltip-trigger[data-tooltip-key])
       =========================== */
function ensureContractTooltipTriggers() {
  const resultsLabels = Array.from(document.querySelectorAll(".results-indicator-label"));
  resultsLabels.forEach((el) => {
    if (!el.classList.contains("tooltip-trigger")) el.classList.add("tooltip-trigger");
  });

  const nationalLabels = Array.from(document.querySelectorAll(".national-indicator-label"));
  nationalLabels.forEach((el) => {
    if (!el.classList.contains("tooltip-trigger")) el.classList.add("tooltip-trigger");
  });

  const requiredIdKeyPairs = [
    ["optout-alt-info", "opt_out_alternative"],
    ["cost-components-info", "cost_components"],
    ["opp-cost-info", "opportunity_cost"],
    ["preference-model-info", "preference_model"],

    ["result-endorsement-info", "result_endorsement"],
    ["result-optout-info", "result_optout"],
    ["result-wtp-trainee-info", "result_wtp_per_trainee"],
    ["result-wtp-cohort-info", "result_wtp_per_cohort"],
    ["result-prog-cost-info", "result_cost_programme"],
    ["result-total-cost-info", "result_cost_total"],
    ["result-net-benefit-info", "result_npv"],
    ["result-bcr-info", "result_bcr"],
    ["result-epi-graduates-info", "result_epi_graduates"],
    ["result-epi-outbreaks-info", "result_epi_outbreaks"],
    ["result-epi-benefit-info", "result_epi_benefit"],

    ["natsim-total-cost-info", "national_total_cost"],
    ["natsim-total-benefit-info", "national_total_benefit"],
    ["natsim-net-benefit-info", "national_net_benefit"],
    ["natsim-bcr-info", "national_bcr"],
    ["natsim-graduates-info", "national_graduates"],
    ["natsim-outbreaks-info", "national_outbreaks"],
    ["natsim-total-wtp-info", "national_total_wtp"]
  ];

  requiredIdKeyPairs.forEach(([id, key]) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.add("tooltip-trigger");
    el.setAttribute("data-tooltip-key", key);
    el.removeAttribute("title");
  });

  const legacy = Array.from(document.querySelectorAll("[data-tooltip]"));
  legacy.forEach((el) => {
    if (el.hasAttribute("data-tooltip-key")) return;
    el.classList.add("tooltip-trigger");
  });
}

function initTooltips() {
  const tooltipEl = document.getElementById("globalTooltip");
  if (!tooltipEl) return;

  const titleEl = tooltipEl.querySelector(".tooltip-title");
  const bodyEl = tooltipEl.querySelector(".tooltip-body");
  if (!titleEl || !bodyEl) return;

  appState.tooltip.tooltipEl = tooltipEl;
  appState.tooltip.titleEl = titleEl;
  appState.tooltip.bodyEl = bodyEl;

  tooltipEl.setAttribute("role", "tooltip");
  tooltipEl.style.position = tooltipEl.style.position || "absolute";
  tooltipEl.style.zIndex = tooltipEl.style.zIndex || "9999";
  tooltipEl.style.visibility = "hidden";
  tooltipEl.style.opacity = "0";
  tooltipEl.style.pointerEvents = "none";

  ensureContractTooltipTriggers();

  function getTooltipPayload(target) {
    const key = target.getAttribute("data-tooltip-key");
    if (key && TOOLTIP_LIBRARY[key]) return TOOLTIP_LIBRARY[key];

    if (key && key.startsWith("result_") && !TOOLTIP_LIBRARY[key]) {
      return {
        title: "Indicator",
        body:
          "This indicator summarises the results shown. See the settings section for the assumptions used. In general, ratios greater than 1 and net benefits greater than 0 mean the benefits outweigh the costs."
      };
    }
    if (key && key.startsWith("national_") && !TOOLTIP_LIBRARY[key]) {
      return {
        title: "Indicator",
        body:
          "This indicator summarises a national scale output derived by aggregating cohort level results. It depends on the configured cohorts, trainees per cohort, endorsement share, and the epidemiological and economic assumptions in settings."
      };
    }

    const legacyText =
      target.getAttribute("data-tooltip") ||
      target.getAttribute("aria-label") ||
      "";

    if (legacyText) {
      return {
        title: "",
        body: legacyText
      };
    }

    return null;
  }

  function setTooltipContent(payload) {
    titleEl.textContent = safeText(payload.title || "");
    bodyEl.textContent = safeText(payload.body || "");
  }

  function positionTooltip(target) {
    const rect = target.getBoundingClientRect();
    const margin = 10;
    const offset = 10;

    tooltipEl.style.left = "0px";
    tooltipEl.style.top = "0px";
    tooltipEl.style.visibility = "hidden";
    tooltipEl.style.opacity = "0";

    tooltipEl.style.pointerEvents = "none";

    tooltipEl.style.visibility = "hidden";
    tooltipEl.style.opacity = "0";
    tooltipEl.style.display = "block";

    const tipRect = tooltipEl.getBoundingClientRect();
    const viewportW = window.innerWidth;
    const viewportH = window.innerHeight;

    let top = rect.top - tipRect.height - offset;
    let left = rect.left + rect.width / 2 - tipRect.width / 2;

    if (top < margin) {
      top = rect.bottom + offset;
    }

    left = clamp(left, margin, viewportW - tipRect.width - margin);

    if (top + tipRect.height > viewportH - margin) {
      const altTop = rect.top - tipRect.height - offset;
      if (altTop >= margin) top = altTop;
      top = clamp(top, margin, viewportH - tipRect.height - margin);
    }

    tooltipEl.style.left = `${left + window.scrollX}px`;
    tooltipEl.style.top = `${top + window.scrollY}px`;
    tooltipEl.style.visibility = "visible";
    tooltipEl.style.opacity = "1";
  }

  function showTooltip(target) {
    const payload = getTooltipPayload(target);
    if (!payload) return;

    if (appState.tooltip.hideTimeout) {
      clearTimeout(appState.tooltip.hideTimeout);
      appState.tooltip.hideTimeout = null;
    }

    appState.tooltip.currentTarget = target;
    setTooltipContent(payload);
    positionTooltip(target);
  }

  function hideTooltip() {
    appState.tooltip.currentTarget = null;
    tooltipEl.style.opacity = "0";
    tooltipEl.style.visibility = "hidden";
  }

  function onEnter(target) {
    showTooltip(target);
  }

  function onLeave(target, related) {
    if (!appState.tooltip.currentTarget) return;
    if (related && (target === related || target.contains(related))) return;
    appState.tooltip.hideTimeout = setTimeout(hideTooltip, 90);
  }

  const triggers = Array.from(document.querySelectorAll(".tooltip-trigger"));

  triggers.forEach((el) => {
    el.addEventListener("mouseenter", () => onEnter(el));
    el.addEventListener("mouseleave", (e) => onLeave(el, e.relatedTarget));
    el.addEventListener("focusin", () => onEnter(el));
    el.addEventListener("focusout", () => onLeave(el, null));
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") hideTooltip();
  });

  window.addEventListener("scroll", () => {
    if (appState.tooltip.currentTarget) {
      positionTooltip(appState.tooltip.currentTarget);
    }
  });

  window.addEventListener("resize", () => {
    if (appState.tooltip.currentTarget) {
      positionTooltip(appState.tooltip.currentTarget);
    }
  });
}

/* ===========================
   Definitions for WTP, mixed logit and key sections
   =========================== */

function initDefinitionTooltips() {
  const wtpInfo = document.getElementById("wtp-info");
  if (wtpInfo) {
    wtpInfo.classList.add("tooltip-trigger");
    if (!wtpInfo.getAttribute("data-tooltip-key")) {
      wtpInfo.setAttribute("data-tooltip", "WTP per trainee per month is derived from the preference model by dividing attribute coefficients by the cost coefficient. It is an approximate rupee value stakeholders attach to this configuration. Total WTP aggregates this value across trainees and cohorts. All benefit values are indicative approximations.");
    }
    wtpInfo.removeAttribute("title");
  }

  const wtpSectionInfo = document.getElementById("wtp-section-info");
  if (wtpSectionInfo) {
    wtpSectionInfo.classList.add("tooltip-trigger");
    if (!wtpSectionInfo.getAttribute("data-tooltip-key")) {
      wtpSectionInfo.setAttribute(
        "data-tooltip",
        "WTP indicators summarise how much value stakeholders attach to each configuration in rupees per trainee and over all cohorts. They are based on the mixed logit preference model and should be read as indicative support rather than precise market prices."
      );
    }
    wtpSectionInfo.removeAttribute("title");
  }

  const mxlInfo = document.getElementById("mixedlogit-info");
  if (mxlInfo) {
    mxlInfo.classList.add("tooltip-trigger");
    if (!mxlInfo.getAttribute("data-tooltip-key")) {
      mxlInfo.setAttribute(
        "data-tooltip",
        "The mixed logit preference model allows preferences to vary across decision makers instead of assuming a single average pattern, which makes endorsement and WTP estimates more flexible."
      );
    }
    mxlInfo.removeAttribute("title");
  }

  const epiInfo = document.getElementById("epi-implications-info");
  if (epiInfo) {
    epiInfo.classList.add("tooltip-trigger");
    if (!epiInfo.getAttribute("data-tooltip-key")) {
      epiInfo.setAttribute(
        "data-tooltip",
        "Graduates and outbreak responses are obtained by combining endorsement with cohort size and number of cohorts. The indicative outbreak cost saving per cohort converts expected outbreak responses into monetary terms using the outbreak value and planning horizon set in the settings."
      );
    }
    epiInfo.removeAttribute("title");
  }

  const endorseInfo = document.getElementById("endorsement-optout-info");
  if (endorseInfo) {
    endorseInfo.classList.add("tooltip-trigger");
    if (!endorseInfo.getAttribute("data-tooltip-key")) {
      endorseInfo.setAttribute(
        "data-tooltip",
        "These percentages come from the mixed logit preference model and show how attractive the configuration is relative to opting out in the preference study."
      );
    }
    endorseInfo.removeAttribute("title");
  }

  const sensInfo = document.getElementById("sensitivity-headline-info");
  if (sensInfo) {
    sensInfo.classList.add("tooltip-trigger");
    if (!sensInfo.getAttribute("data-tooltip-key")) {
      sensInfo.setAttribute(
        "data-tooltip",
        "In this summary, the cost column shows the economic cost for each scenario over the selected time horizon. Total economic cost and net benefit are aggregated across all cohorts in millions of rupees. Total WTP benefits summarise how much value stakeholders place on each configuration, while the outbreak response column isolates the part of that value linked to faster detection and response. Epidemiological outbreak benefits appear when the outbreak benefit switch is on and the epidemiological module is active. The effective WTP benefit scales total WTP by the endorsement rate used in the calculation. Benefit cost ratios compare total benefits with total costs, and net present values show the difference between benefits and costs in rupee terms. Values above one for benefit cost ratios and positive net present values indicate that estimated benefits exceed costs under the current assumptions."
      );
    }
    sensInfo.removeAttribute("title");
  }

  const copilotInfo = document.getElementById("copilot-howto-info");
  const copilotText = document.getElementById("copilot-howto-text");
  if (copilotInfo) {
    copilotInfo.classList.add("tooltip-trigger");
    if (!copilotInfo.getAttribute("data-tooltip-key")) {
      copilotInfo.setAttribute(
        "data-tooltip",
        "First, use the other STEPS tabs to define a scenario you want to interpret. Apply the configuration, review endorsement, WTP, costs and epidemiological outbreak benefits, and check the national and sensitivity views. When you are ready, move to the Copilot tab to prepare a narrative briefing. When you press the Copilot button, STEPS rebuilds the interpretation prompt using the latest scenario and model outputs. The prompt combines a short description of STEPS, instructions for Copilot and the full JSON export for the current scenario. The aim is to guide Copilot to prepare a three to five page policy brief for discussions with ministries, World Bank staff and other partners. The brief is requested as a narrative report with clear sections such as background, scenario description, endorsement patterns, costs, epidemiological benefits, benefit cost ratios, net present values, distributional considerations and implementation notes, and includes compact tables for key indicators. After copying the text from the prompt panel, open Microsoft Copilot in a new browser tab or in the window that STEPS opens, paste the full content into the prompt box and run it. You can then edit the draft policy brief in Copilot or in your preferred word processor, keeping a record of the assumptions and JSON values supplied by STEPS."
      );
    }
    copilotInfo.removeAttribute("title");
  }
  if (copilotText) {
    copilotText.textContent =
      "Define a scenario in the other tabs, then use this Copilot tab to generate a draft policy brief. Copy the prepared prompt into Microsoft Copilot and refine the brief there.";
  }
}

/* ===========================
   Tabs (UI contract: .tablink, .tabcontent, data-tab points to panel id)
   =========================== */

function initTabs() {
  const tabLinks = Array.from(document.querySelectorAll(".tablink[data-tab]"));
  const panels = Array.from(document.querySelectorAll(".tabcontent"));

  if (!tabLinks.length || !panels.length) return;

  panels.forEach((p) => {
    if (!p.dataset.defaultDisplay) {
      const computed = window.getComputedStyle(p).display;
      p.dataset.defaultDisplay = computed && computed !== "none" ? computed : "block";
    }
  });

  function openTab(panelId, btn) {
    panels.forEach((panel) => {
      panel.style.display = "none";
      panel.classList.remove("active");
    });
    tabLinks.forEach((b) => {
      b.classList.remove("active");
      b.setAttribute("aria-selected", "false");
    });

    const panel = document.getElementById(panelId);
    if (panel) {
      panel.style.display = panel.dataset.defaultDisplay || "block";
      panel.classList.add("active");
    }
    if (btn) {
      btn.classList.add("active");
      btn.setAttribute("aria-selected", "true");
    }
  }

  tabLinks.forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = btn.getAttribute("data-tab");
      if (!target) return;
      openTab(target, btn);
    });
  });

  const currentActive = tabLinks.find((b) => b.classList.contains("active"));
  const initial = currentActive || tabLinks[0];
  if (initial) {
    const target = initial.getAttribute("data-tab");
    if (target) openTab(target, initial);
  }
}

/* ===========================
   Guided tour
   =========================== */

function initGuidedTour() {
  const trigger = document.getElementById("btn-start-tour");
  if (!trigger) return;

  const steps = Array.from(document.querySelectorAll("[data-tour-step]"));
  if (!steps.length) return;

  appState.tour.steps = steps;

  const overlay = document.createElement("div");
  overlay.id = "tour-overlay";
  overlay.className = "tour-overlay hidden";

  const popover = document.createElement("div");
  popover.id = "tour-popover";
  popover.className = "tour-popover hidden";
  popover.innerHTML = `
    <div class="tour-popover-header">
      <h3 id="tour-title"></h3>
      <button class="tour-close-btn" type="button" aria-label="Close tour">×</button>
    </div>
    <div class="tour-popover-body" id="tour-body"></div>
    <div class="tour-popover-footer">
      <span class="tour-step-indicator" id="tour-indicator"></span>
      <div class="tour-buttons">
        <button type="button" class="btn-ghost-small" id="tour-prev">Previous</button>
        <button type="button" class="btn-primary-small" id="tour-next">Next</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);
  document.body.appendChild(popover);

  appState.tour.overlayEl = overlay;
  appState.tour.popoverEl = popover;

  function endTour() {
    overlay.classList.add("hidden");
    popover.classList.add("hidden");
  }

  function showStep(index) {
    const stepsArr = appState.tour.steps;
    if (!stepsArr.length) return;
    const i = clamp(index, 0, stepsArr.length - 1);
    appState.tour.currentIndex = i;
    const el = stepsArr[i];
    if (!el) return;

    const title = el.getAttribute("data-tour-title") || "STEPS tour";
    const content = el.getAttribute("data-tour-content") || "";

    document.getElementById("tour-title").textContent = title;
    document.getElementById("tour-body").textContent = content;
    document.getElementById("tour-indicator").textContent = `Step ${i + 1} of ${stepsArr.length}`;

    overlay.classList.remove("hidden");
    popover.classList.remove("hidden");

    positionTourPopover(popover, el);
  }

  function positionTourPopover(popoverEl, targetEl) {
    const rect = targetEl.getBoundingClientRect();
    const popRect = popoverEl.getBoundingClientRect();
    let top = rect.bottom + 8 + window.scrollY;
    let left = rect.left + window.scrollX + rect.width / 2 - popRect.width / 2;

    if (left < 8) left = 8;
    if (left + popRect.width > window.scrollX + window.innerWidth - 8) {
      left = window.scrollX + window.innerWidth - popRect.width - 8;
    }
    if (top + popRect.height > window.scrollY + window.innerHeight - 8) {
      top = rect.top + window.scrollY - popRect.height - 10;
    }

    popoverEl.style.left = `${left}px`;
    popoverEl.style.top = `${top}px`;
  }

  trigger.addEventListener("click", () => showStep(0));

  overlay.addEventListener("click", endTour);
  popover.querySelector(".tour-close-btn").addEventListener("click", endTour);
  popover.querySelector("#tour-prev").addEventListener("click", () => showStep(appState.tour.currentIndex - 1));
  popover.querySelector("#tour-next").addEventListener("click", () => {
    if (appState.tour.currentIndex >= appState.tour.steps.length - 1) endTour();
    else showStep(appState.tour.currentIndex + 1);
  });

  window.addEventListener("resize", () => {
    if (!overlay.classList.contains("hidden") && appState.tour.steps.length) {
      const el = appState.tour.steps[appState.tour.currentIndex];
      if (el) positionTourPopover(popover, el);
    }
  });

  window.addEventListener("scroll", () => {
    if (!overlay.classList.contains("hidden") && appState.tour.steps.length) {
      const el = appState.tour.steps[appState.tour.currentIndex];
      if (el) positionTourPopover(popover, el);
    }
  });
}

/* ===========================
   Configuration and results
   =========================== */

function labelForTier(tier) {
  if (tier === "frontline") return "Frontline";
  if (tier === "intermediate") return "Intermediate";
  if (tier === "advanced") return "Advanced";
  return safeText(tier || "");
}

function labelForCareer(career) {
  if (career === "certificate") return "Certificate";
  if (career === "uniqual") return "University qualification";
  if (career === "career_path") return "Career pathway";
  return safeText(career || "");
}

function labelForMentorship(m) {
  if (m === "low") return "Low";
  if (m === "medium") return "Medium";
  if (m === "high") return "High";
  return safeText(m || "");
}

function labelForDelivery(d) {
  if (d === "blended") return "Blended";
  if (d === "inperson") return "In-person";
  if (d === "online") return "Online";
  return safeText(d || "");
}

/**
 * Build an informative default scenario name that mirrors the current configuration.
 * This does not change any underlying computations; it is a labelling convenience only.
 */
function buildDefaultScenarioName(fields) {
  const tier = labelForTier(fields.tier);
  const mentorship = labelForMentorship(fields.mentorship);
  const delivery = labelForDelivery(fields.delivery);
  const career = labelForCareer(fields.career);

  const cohorts = Number(fields.cohorts) || 0;
  const trainees = Number(fields.traineesPerCohort) || 0;
  const cost = Number(fields.costPerTraineePerMonth) || 0;

  const scale = cohorts > 0 && trainees > 0 ? `${formatNumber(cohorts, 0)} cohorts × ${formatNumber(trainees, 0)} trainees` : "";
  const costPart = cost > 0 ? `₹${formatNumber(cost, 0)}/mo` : "";

  const parts = [
    `${tier}`,
    `${mentorship} mentorship`,
    delivery,
    career,
    scale,
    costPart
  ].filter((p) => String(p || "").trim().length > 0);

  return parts.join(" | ");
}



function getPlanningHorizonFromInputs() {
  let planningHorizonYears = appState.epiSettings.general.planningHorizonYears;
  const candidates = [
    document.getElementById("planning-horizon-config"),
    document.getElementById("adv-planning-horizon"),
    document.getElementById("planning-horizon")
  ];
  for (const el of candidates) {
    if (el) {
      const v = Number(el.value);
      if (!isNaN(v) && v > 0) {
        planningHorizonYears = v;
        break;
      }
    }
  }
  appState.epiSettings.general.planningHorizonYears = planningHorizonYears;
  return planningHorizonYears;
}

function getConfigFromForm() {
  const tier = document.getElementById("program-tier").value;
  const career = document.getElementById("career-track").value;
  const mentorship = document.getElementById("mentorship").value;
  const delivery = document.getElementById("delivery").value;

  let response = "7";
  const responseEl = document.getElementById("response");
  if (responseEl) {
    responseEl.value = "7";
    response = "7";
  }

  const costSlider = Number(document.getElementById("cost-slider").value);
  const trainees = Number(document.getElementById("trainees").value);
  // cohorts value read after capacity constraints

  // Settings: planning horizon is stored in appState.epiSettings.general
  let planningHorizonYears = getPlanningHorizonFromInputs();

  // STEPS UPGRADE: cohorts are constrained by tier duration, horizon and available training sites/hubs
  enforceCohortCapacityConstraints({ tierKey: tier, clampIfNeeded: true, source: "config" });

  const cohorts = Number(document.getElementById("cohorts").value);

  const oppIncluded = document.getElementById("opp-toggle").classList.contains("on");

  
  const scenarioNameEl = document.getElementById("scenario-name");
  const rawName = scenarioNameEl ? (scenarioNameEl.value || "").trim() : "";
  const signature = JSON.stringify({
    tier,
    career,
    mentorship,
    delivery,
    cohorts,
    traineesPerCohort: trainees,
    costPerTraineePerMonth: costSlider
  });

  // Auto-name behaviour:
  // - If the user has not manually edited the name (auto mode), refresh it whenever key fields change.
  // - If the name is empty, also auto-name.
  const shouldAutoName = !!appState.autoScenarioName || !rawName;

  let scenarioName = rawName;
  if (shouldAutoName) {
    if (signature !== appState._lastAutoNameSignature || !rawName) {
      scenarioName = buildDefaultScenarioName({
        tier,
        career,
        mentorship,
        delivery,
        cohorts,
        traineesPerCohort: trainees,
        costPerTraineePerMonth: costSlider
      });
      appState._lastAutoNameSignature = signature;
      if (scenarioNameEl) scenarioNameEl.value = scenarioName;
    } else {
      scenarioName = rawName || buildDefaultScenarioName({
        tier,
        career,
        mentorship,
        delivery,
        cohorts,
        traineesPerCohort: trainees,
        costPerTraineePerMonth: costSlider
      });
    }
  }

  const scenarioNotesEl = document.getElementById("scenario-notes");
  const scenarioNotes = scenarioNotesEl ? (scenarioNotesEl.value || "").trim() : "";

  // NEW: mentor support cost per cohort (base) and feasibility inputs (defensive reads)
  const mentorBaseEl = document.getElementById("mentor-support-cost-per-cohort");
  let mentorSupportCostPerCohortBase = 0;
  if (mentorBaseEl) {
    const v = Number(mentorBaseEl.value);
    if (!isNaN(v) && v >= 0) mentorSupportCostPerCohortBase = v;
  }

  const availableMentorsEl = document.getElementById("available-mentors-national");
  let availableMentorsNational = 200;
  if (availableMentorsEl) {
    const v = Number(availableMentorsEl.value);
    if (!isNaN(v) && v >= 0) availableMentorsNational = v;
  }

  const sitesEl = document.getElementById("available-training-sites");
  let availableTrainingSites = 0;
  if (sitesEl) {
    const v = Number(sitesEl.value);
    if (!isNaN(v) && v >= 0) availableTrainingSites = v;
  }

  const maxCohortsEl = document.getElementById("max-cohorts-per-site");
  let maxCohortsPerSitePerYear = 0;
  if (maxCohortsEl) {
    const v = Number(maxCohortsEl.value);
    if (!isNaN(v) && v >= 0) maxCohortsPerSitePerYear = v;
  }

  const crossSectorEl = document.getElementById("cross-sector-multiplier");
  let crossSectorBenefitMultiplier = 1.0;
  if (crossSectorEl) {
    const v = Number(crossSectorEl.value);
    if (!isNaN(v) && v > 0) crossSectorBenefitMultiplier = v;
  }
  // Clamp recommended range if provided
  crossSectorBenefitMultiplier = Math.max(0.8, Math.min(2.0, crossSectorBenefitMultiplier));

  // NEW: optional export notes (persisted into scenarios for briefing/exports)
  const enablersEl = document.getElementById("export-enablers");
  const risksEl = document.getElementById("export-risks");
  const exportEnablers = enablersEl ? (enablersEl.value || "").trim() : "";
  const exportRisks = risksEl ? (risksEl.value || "").trim() : "";

  // Salary-based opportunity cost inputs (stored on scenario config for persistence)
  const salaryInputs = readSalaryInputsFromDom();
  const salaryMonthlyFaculty = salaryInputs.facultyMonthly;
  const salaryMonthlyCoordinator = salaryInputs.coordinatorMonthly;
  const salaryMonthlyParticipant = salaryInputs.participantMonthly;

  return {
    tier,
    career,
    mentorship,
    delivery,
    response,
    costPerTraineePerMonth: costSlider,
    traineesPerCohort: trainees,
    cohorts,
    planningHorizonYears,
    opportunityCostIncluded: oppIncluded,
    name: scenarioName,
    notes: scenarioNotes,
    preferenceModel: "Mixed logit model from the preference study",

    // NEW fields
    mentorSupportCostPerCohortBase,
    availableMentorsNational,
    availableTrainingSites,
    maxCohortsPerSitePerYear,
    crossSectorBenefitMultiplier,
    exportEnablers,
    exportRisks,

    // Salary-based opportunity cost inputs
    salaryMonthlyFaculty,
    salaryMonthlyCoordinator,
    salaryMonthlyParticipant
  };
}

// Backwards-compatible alias used by some UI handlers
function getConfigFromUI() {
  return getConfigFromForm();
}



function tierEffect(tier) {
  return MXL_COEFS.tier[tier] || 0;
}

function careerEffect(career) {
  return MXL_COEFS.career[career] || 0;
}

function mentorshipEffect(m) {
  return MXL_COEFS.mentorship[m] || 0;
}

function deliveryEffect(d) {
  return MXL_COEFS.delivery[d] || 0;
}

function responseEffect(r) {
  return MXL_COEFS.response[r] || 0;
}

function computeEndorsementAndWTP(config) {
  const costThousands = config.costPerTraineePerMonth / 1000;
  const utilProgram =
    MXL_COEFS.ascProgram +
    tierEffect(config.tier) +
    careerEffect(config.career) +
    mentorshipEffect(config.mentorship) +
    deliveryEffect(config.delivery) +
    responseEffect(config.response) +
    MXL_COEFS.costPerThousand * costThousands;

  const utilOptOut = MXL_COEFS.ascOptOut;

  const maxU = Math.max(utilProgram, utilOptOut);
  const expProg = Math.exp(utilProgram - maxU);
  const expOpt = Math.exp(utilOptOut - maxU);
  const denom = expProg + expOpt;

  const endorseProb = denom > 0 ? expProg / denom : 0.5;
  const optOutProb = 1 - endorseProb;

  const nonCostUtility =
    MXL_COEFS.ascProgram +
    tierEffect(config.tier) +
    careerEffect(config.career) +
    mentorshipEffect(config.mentorship) +
    deliveryEffect(config.delivery) +
    responseEffect(config.response);

  const wtpPerTraineePerMonth = (nonCostUtility / Math.abs(MXL_COEFS.costPerThousand)) * 1000;

  return {
    endorseRate: clamp(endorseProb * 100, 0, 100),
    optOutRate: clamp(optOutProb * 100, 0, 100),
    wtpPerTraineePerMonth
  };
}

function mentorshipMultiplier(intensity) {
  if (String(intensity) === "medium") return 1.3;
  if (String(intensity) === "high") return 1.7;
  return 1.0;
}


function computeSalaryBasedOpportunityCostPerCohort(config) {
  // Always compute the raw salary-based opportunity cost for transparency.
  const tier = config.tier || "intermediate";
  const months = TIER_MONTHS[tier] || 12;
  const totalDays = months * SALARY_OC_DAYS_PER_MONTH;

  const delivery = String(config.delivery || "").toLowerCase();
  const isInPerson = delivery === "inperson";
  const isBlended = delivery === "blended";
  const isOnline = delivery === "online";

  // Contact days assumptions apply to blended (and online as Dc=0).
  const contactDaysAssumed = SALARY_OC_CONTACT_DAYS_BY_TIER[tier] ?? 0;
  const contactDays = isBlended ? contactDaysAssumed : (isOnline ? 0 : totalDays);

  const salaryFaculty = Number(config.salaryMonthlyFaculty ?? SALARY_OC_DEFAULTS.facultyMonthly);
  const salaryCoordinator = Number(config.salaryMonthlyCoordinator ?? SALARY_OC_DEFAULTS.coordinatorMonthly);
  const salaryParticipant = Number(config.salaryMonthlyParticipant ?? SALARY_OC_DEFAULTS.participantMonthly);

  const facultyMonthly = isFinite(salaryFaculty) && salaryFaculty >= 0 ? salaryFaculty : SALARY_OC_DEFAULTS.facultyMonthly;
  const coordinatorMonthly = isFinite(salaryCoordinator) && salaryCoordinator >= 0 ? salaryCoordinator : SALARY_OC_DEFAULTS.coordinatorMonthly;
  const participantMonthly = isFinite(salaryParticipant) && salaryParticipant >= 0 ? salaryParticipant : SALARY_OC_DEFAULTS.participantMonthly;

  const nParticipant = Number(config.traineesPerCohort || 0);
  const nCoordinator = 1;

  // Faculty proxy: mentors required per cohort (ties faculty time to mentorship intensity and cohort size)
  let nFaculty = 0;
  try {
    const cap = computeCapacity({ ...config, cohorts: 1 });
    nFaculty = Number(cap && cap.mentorsPerCohort) || 0;
  } catch (e) {
    nFaculty = 0;
  }

  // Fully in-person: salary counted for full programme duration (months).
  // Blended: participants/coordinators only contact days; faculty contact days + 50% of non-contact duration.
  // Online: treated as blended with Dc=0.
  let ocFaculty = 0;
  let ocCoordinator = 0;
  let ocParticipant = 0;

  if (isInPerson) {
    ocFaculty = facultyMonthly * months * nFaculty;
    ocCoordinator = coordinatorMonthly * months * nCoordinator;
    ocParticipant = participantMonthly * months * nParticipant;
  } else {
    const dcMonths = contactDays / SALARY_OC_DAYS_PER_MONTH;
    const nonContactMonths = Math.max(0, (totalDays - contactDays)) / SALARY_OC_DAYS_PER_MONTH;

    ocCoordinator = coordinatorMonthly * dcMonths * nCoordinator;
    ocParticipant = participantMonthly * dcMonths * nParticipant;

    ocFaculty = facultyMonthly * (dcMonths + 0.5 * nonContactMonths) * nFaculty;
  }

  const total = ocFaculty + ocCoordinator + ocParticipant;

  return {
    byRole: {
      faculty: ocFaculty,
      coordinator: ocCoordinator,
      participant: ocParticipant
    },
    total,
    meta: {
      tier,
      delivery,
      months,
      totalDays,
      contactDays,
      contactDaysAssumed,
      nFaculty,
      nCoordinator,
      nParticipant,
      salaryMonthly: {
        faculty: facultyMonthly,
        coordinator: coordinatorMonthly,
        participant: participantMonthly
      }
    }
  };
}

function computeCosts(config) {
  const months = TIER_MONTHS[config.tier] || 12;
  const directCostPerTraineePerMonth = config.costPerTraineePerMonth;
  const trainees = config.traineesPerCohort;

  const programmeCostPerCohort = directCostPerTraineePerMonth * months * trainees;

  // NEW: mentor support costs (base × mentorship multiplier)
  const mentorBase = Number(config.mentorSupportCostPerCohortBase || 0);
  const mentorMult = mentorshipMultiplier(config.mentorship);
  const mentorCostPerCohort = mentorBase * mentorMult;

  // Direct cost excludes opportunity cost, but includes mentor support costs
  const directCostPerCohort = programmeCostPerCohort + mentorCostPerCohort;

  const templatesForTier = COST_TEMPLATES[config.tier];
  const template =
    (COST_CONFIG && COST_CONFIG[config.tier] && COST_CONFIG[config.tier].combined) ||
    (templatesForTier && templatesForTier.combined);

  let oppRate = template ? template.oppRate : 0;
  if (!config.opportunityCostIncluded) {
    oppRate = 0;
  }

  // Keep prior behaviour: opportunity cost is applied to programme delivery costs
  const opportunityCost = programmeCostPerCohort * oppRate;

  // Salary-based opportunity cost (additional module)
  const salaryOppRaw = computeSalaryBasedOpportunityCostPerCohort(config);
  const salaryBasedOpportunityCostIncluded = config.opportunityCostIncluded ? (salaryOppRaw.total || 0) : 0;

  const totalEconomicCost = directCostPerCohort + opportunityCost + salaryBasedOpportunityCostIncluded;

  // Convenience totals across all cohorts
  const totalMentorCostAllCohorts = mentorCostPerCohort * (config.cohorts || 0);
  const totalDirectCostAllCohorts = directCostPerCohort * (config.cohorts || 0);
  const totalEconomicCostAllCohorts = totalEconomicCost * (config.cohorts || 0);

  
  // Internal reconciliation check for accuracy
  try {
    const recon = directCostPerCohort + (opportunityCostIncluded ? opportunityCost : 0) + (opportunityCostIncluded ? salaryBasedOpportunityCostIncluded : 0);
    const diff = Math.abs(recon - totalEconomicCost);
    if (diff > 0.5) {
      console.warn("[STEPS] Cost reconciliation difference detected (per cohort).", { recon, totalEconomicCost, diff, tier: config.tier, delivery: config.delivery });
    }
  } catch (e) {
    // ignore
  }

return {
    programmeCostPerCohort,
    mentorSupportCostPerCohortBase: mentorBase,
    mentorCostMultiplier: mentorMult,
    mentorCostPerCohort,
    directCostPerCohort,
    opportunityCostPerCohort: opportunityCost,
    salaryBasedOpportunityCostPerCohort: salaryBasedOpportunityCostIncluded,
    salaryBasedOpportunityCostPerCohortRaw: salaryOppRaw.total || 0,
    salaryBasedOpportunityCostByRole: salaryOppRaw.byRole,
    salaryBasedOpportunityCostMeta: salaryOppRaw.meta,
    totalEconomicCostPerCohort: totalEconomicCost,
    totalMentorCostAllCohorts,
    totalDirectCostAllCohorts,
    totalEconomicCostAllCohorts,
    template
  };
}


function computeEpidemiological(config, endorseRate) {
  const tierSettings = appState.epiSettings.tiers[config.tier];
  const general = appState.epiSettings.general;

  
let completionRate = tierSettings.completionRate;
// Optional per-config completion override (0-1 or 0-100). Used by Baseline editor.
if (config && config.completionRateOverride !== undefined && config.completionRateOverride !== null && config.completionRateOverride !== "") {
  const raw = Number(config.completionRateOverride);
  if (isFinite(raw)) {
    completionRate = raw > 1 ? raw / 100 : raw;
  }
}
completionRate = Math.max(0, Math.min(1, Number(completionRate)));
  const outbreaksPerGrad = tierSettings.outbreaksPerGraduatePerYear;
  const valuePerOutbreak = tierSettings.valuePerOutbreak;

  // Optional non-outbreak value (per graduate per year); defaults to 0 if not used.
  const valuePerGraduate = Number(tierSettings.valuePerGraduate || 0);

  const planningYears = (Number(config && config.planningHorizonYears) > 0) ? Number(config.planningHorizonYears) : general.planningHorizonYears;
  const discountRate = general.epiDiscountRate;

  const pvFactor = presentValueFactor(discountRate, planningYears);
  const endorseFactor = endorseRate / 100;

  const months = TIER_MONTHS[config.tier] || 12;

  const enrolledPerCohort = config.traineesPerCohort;
  const completedPerCohort = enrolledPerCohort * completionRate;
  const graduatesEffective = completedPerCohort * endorseFactor;

  const graduatesAllCohorts = graduatesEffective * config.cohorts;

  const respMultiplier = RESPONSE_TIME_MULTIPLIERS[String(config.response)] || 1;

  const outbreaksPerYearPerCohort = graduatesEffective * outbreaksPerGrad * respMultiplier;
  const outbreaksPerYearNational = outbreaksPerYearPerCohort * config.cohorts;

  // Base PV benefits
  const graduateAnnualBenefitPerCohort = graduatesEffective * valuePerGraduate;
  const graduateBenefitPerCohort = graduateAnnualBenefitPerCohort * pvFactor;

  const outbreakAnnualBenefitPerCohort = outbreaksPerYearPerCohort * valuePerOutbreak;
  const outbreakPVPerCohort = outbreakAnnualBenefitPerCohort * pvFactor;

  // NEW: cross-sector / One Health multiplier applied to epidemiological benefits only
  let crossSectorMultiplier = Number(config.crossSectorBenefitMultiplier || 1.0);
  if (isNaN(crossSectorMultiplier) || crossSectorMultiplier <= 0) crossSectorMultiplier = 1.0;
  crossSectorMultiplier = Math.max(0.8, Math.min(2.0, crossSectorMultiplier));

  const graduateBenefitAdj = graduateBenefitPerCohort * crossSectorMultiplier;
  const outbreakPVAdj = outbreakPVPerCohort * crossSectorMultiplier;

  const totalEpiBenefitPerCohort = graduateBenefitAdj + outbreakPVAdj;

  return {
    months,
    graduatesPerCohort: graduatesEffective,
    graduatesAllCohorts,
    outbreaksPerYearPerCohort,
    outbreaksPerYearNational,
    epiBenefitPerCohort: totalEpiBenefitPerCohort,
    graduateBenefitPerCohort: graduateBenefitAdj,
    outbreakPVPerCohort: outbreakPVAdj,
    planningYears,
    discountRate,
    completionRate,
    outbreaksPerGraduatePerYear: outbreaksPerGrad,
    valuePerOutbreak,
    valuePerGraduate,
    crossSectorMultiplier
  };
}

function mentorshipMentorCapacity(intensity) {
  if (String(intensity) === "high") return 2;
  if (String(intensity) === "medium") return 3.5;
  return 5;
}

function computeCapacity(config) {
  const trainees = Number(config.traineesPerCohort || 0);
  const cohorts = Number(config.cohorts || 0);

  const fellowsPerMentor = mentorshipMentorCapacity(config.mentorship);
  const mentorsPerCohort = fellowsPerMentor > 0 ? Math.ceil(trainees / fellowsPerMentor) : 0;
  const totalMentorsRequired = mentorsPerCohort * cohorts;

  const availableMentors = Number(config.availableMentorsNational ?? 200);
  const mentorShortfall = Math.max(0, totalMentorsRequired - (isNaN(availableMentors) ? 0 : availableMentors));
  const withinMentorCapacity = totalMentorsRequired <= (isNaN(availableMentors) ? 0 : availableMentors);

  // Optional sites/hubs capacity (if provided)
  const sites = Number(config.availableTrainingSites || 0);
  const maxCohortsPerSite = Number(config.maxCohortsPerSitePerYear || 0);

  let siteCapacity = null;
  let siteGap = null;
  let withinSiteCapacity = null;
  if (!isNaN(sites) && !isNaN(maxCohortsPerSite) && sites > 0 && maxCohortsPerSite > 0) {
    siteCapacity = sites * maxCohortsPerSite;
    siteGap = Math.max(0, cohorts - siteCapacity);
    withinSiteCapacity = cohorts <= siteCapacity;
  }

  const status = withinMentorCapacity ? "Within current capacity" : "Requires capacity expansion";

  return {
    fellowsPerMentor,
    mentorsPerCohort,
    totalMentorsRequired,
    availableMentors: isNaN(availableMentors) ? 0 : availableMentors,
    mentorShortfall,
    status,
    siteCapacity,
    siteGap,
    withinSiteCapacity
  };
}

function buildAssumptionsForScenario(scenario) {
  const c = scenario.config;
  const tierSettings = appState.epiSettings.tiers[c.tier] || {};
  const general = appState.epiSettings.general || {};

  return {
    planningHorizonYears: scenario.planningYears ?? general.planningHorizonYears,
    discountRate: scenario.discountRate ?? general.epiDiscountRate,
    completionRate: scenario.epi?.completionRate ?? tierSettings.completionRate,
    outbreaksPerGraduatePerYear: scenario.epi?.outbreaksPerGraduatePerYear ?? tierSettings.outbreaksPerGraduatePerYear,
    valuePerOutbreak: scenario.epi?.valuePerOutbreak ?? tierSettings.valuePerOutbreak,
    valuePerGraduate: scenario.epi?.valuePerGraduate ?? tierSettings.valuePerGraduate ?? 0,
    opportunityCostIncluded: !!c.opportunityCostIncluded,
    mentorSupportCostPerCohortBase: Number(c.mentorSupportCostPerCohortBase || 0),
    mentorMultiplierApplied: mentorshipMultiplier(c.mentorship),
    crossSectorBenefitMultiplier: Number(c.crossSectorBenefitMultiplier || 1.0),
    availableMentorsNational: Number(c.availableMentorsNational ?? 200),
    availableTrainingSites: (() => { const v = Number(c.availableTrainingSites); return (isFinite(v) && v > 0) ? v : 1; })(),
    maxCohortsPerSitePerYear: Number(c.maxCohortsPerSitePerYear || 0),

    // Salary-based opportunity cost inputs & assumptions (additional module)
    salaryMonthlyFaculty: Number(c.salaryMonthlyFaculty ?? SALARY_OC_DEFAULTS.facultyMonthly),
    salaryMonthlyCoordinator: Number(c.salaryMonthlyCoordinator ?? SALARY_OC_DEFAULTS.coordinatorMonthly),
    salaryMonthlyParticipant: Number(c.salaryMonthlyParticipant ?? SALARY_OC_DEFAULTS.participantMonthly),
    salaryOcDaysPerMonth: SALARY_OC_DAYS_PER_MONTH,
    salaryOcContactDaysFrontline: SALARY_OC_CONTACT_DAYS_BY_TIER.frontline,
    salaryOcContactDaysIntermediate: SALARY_OC_CONTACT_DAYS_BY_TIER.intermediate,
    salaryOcContactDaysAdvanced: SALARY_OC_CONTACT_DAYS_BY_TIER.advanced
  };
}


function buildEnablersAndRisksText(scenario) {
  if (!scenario || !scenario.config) {
    return {
      enablers: "Configuration aligns with NCDC stewardship and provides structured field epidemiology training.",
      risks: "Monitor mentor availability, site readiness and quality assurance as scale increases."
    };
  }

  const c = scenario.config;
  const tierKey = c.tier || "intermediate";
  const endorsement = Number(scenario.endorseRate || 0);
  const natBcr = typeof scenario.natBcr === "number" ? scenario.natBcr : null;
  const netBenefit = typeof scenario.netBenefitAllCohorts === "number" ? scenario.netBenefitAllCohorts : null;
  const cap = scenario.capacity || {};
  const wtpPerCohort = typeof scenario.wtpPerCohort === "number" ? scenario.wtpPerCohort : null;
  const costPerCohort =
    scenario.costs && typeof scenario.costs.totalEconomicCostPerCohort === "number"
      ? scenario.costs.totalEconomicCostPerCohort
      : null;
  const wtpBcr = costPerCohort && wtpPerCohort !== null && costPerCohort > 0 ? wtpPerCohort / costPerCohort : null;

  const enablers = [];
  const risks = [];

  if (tierKey === "frontline") {
    enablers.push("Frontline configuration supports rapid expansion of basic field epidemiology skills at district and state levels.");
  } else if (tierKey === "intermediate") {
    enablers.push("Intermediate configuration builds operational leadership capacity in surveillance and outbreak investigation.");
  } else if (tierKey === "advanced") {
    enablers.push("Advanced configuration develops high-level analytic and leadership capacity for complex outbreaks and national responses.");
  }

  if (endorsement >= 70) {
    enablers.push("High endorsement in the preference study (≥ 70%), indicating strong perceived programme value among stakeholders.");
  } else if (endorsement >= 50) {
    enablers.push("Moderate endorsement (50-69%), suggesting the configuration is acceptable with clear communication and engagement.");
  } else if (endorsement > 0) {
    risks.push("Lower endorsement (< 50%) in the preference study, signalling potential acceptability challenges that require stakeholder engagement.");
  }

  if (wtpBcr !== null) {
    if (wtpBcr >= 1.2) {
      enablers.push("Perceived programme value (willingness to pay) is comfortably above total economic cost (ratio ≥ 1.2).");
    } else if (wtpBcr >= 1.0) {
      enablers.push("Perceived programme value is broadly in line with total economic cost (ratio around 1.0), but fiscal space and competing priorities still need discussion.");
    } else {
      risks.push("Perceived programme value is below total economic cost (ratio < 1.0), implying the configuration may be financially hard to justify without co-financing or redesign.");
    }
  }

  if (natBcr !== null && natBcr >= 1.0) {
    enablers.push("Discounted outbreak-related epidemiological benefits are at least as large as total economic costs (benefit-cost ratio ≥ 1.0).");
  } else if (natBcr !== null && natBcr < 1.0) {
    risks.push("Under current outbreak assumptions, epidemiological benefits are below total economic costs (benefit-cost ratio < 1.0).");
  }

  if (netBenefit !== null && netBenefit > 0) {
    enablers.push("Net benefits over the chosen planning horizon are positive, supporting the case for investment.");
  } else if (netBenefit !== null && netBenefit <= 0) {
    risks.push("Net benefits are close to zero or negative, pointing to the need for phasing, redesign or additional co-benefits.");
  }

  if (cap.status === "Within current capacity" || cap.status === "Within capacity") {
    enablers.push("Required mentors and sites appear to be within stated national capacity, reducing implementation risk.");
  } else if (cap.status) {
    risks.push("Scenario exceeds current mentorship or site capacity, requiring either capacity expansion or a more phased scale-up.");
  }
  if (typeof cap.mentorShortfall === "number" && cap.mentorShortfall > 0) {
    risks.push(`Mentor shortfall of approximately ${cap.mentorShortfall} mentors unless additional supervisors are mobilised or cohorts are adjusted.`);
  }

  if (!enablers.length) {
    enablers.push("Configuration strengthens the national FETP network under NCDC stewardship and can be aligned with phased implementation.");
  }
  if (!risks.length) {
    risks.push("Implementation risks relate mainly to mentor availability, site readiness, financing flows and maintaining quality at scale.");
  }

  return {
    enablers: enablers.join(" "),
    risks: risks.join(" ")
  };
}

function ensureScenarioHasExportNotes(scenario) {
  if (!scenario || !scenario.config) return;
  const c = scenario.config;

  const existingEnablers = (c.exportEnablers || "").trim();
  const existingRisks = (c.exportRisks || "").trim();

  if (existingEnablers && existingRisks) {
    return;
  }

  const generated = buildEnablersAndRisksText(scenario);
  if (!existingEnablers) {
    c.exportEnablers = generated.enablers;
  }
  if (!existingRisks) {
    c.exportRisks = generated.risks;
  }

  scenario.exportEnablers = c.exportEnablers;
  scenario.exportRisks = c.exportRisks;
}

function getSelectedScenarioIds() {
  const checks = Array.from(document.querySelectorAll('#scenario-table tbody input[type="checkbox"][data-scenario-id]'));
  return checks.filter((c) => c.checked).map((c) => c.getAttribute("data-scenario-id"));
}

function getShortlistedOrTopScenarios(limit = 3) {
  // Priority 1: explicit checkbox selections in the saved scenarios table (if present)
  const selected = new Set(getSelectedScenarioIds());
  const saved = appState.savedScenarios.slice();

  if (selected.size > 0) {
    return saved.filter((s) => selected.has(s.id));
  }

  // Priority 2: persisted shortlist flags on scenarios (for users who shortlist via the Top options panel)
  const shortlisted = saved.filter((s) => !!s.shortlisted);
  if (shortlisted.length > 0) {
    return shortlisted.slice(0, Math.min(limit, shortlisted.length));
  }

  // Default: top scenarios by net benefit (all cohorts), descending
  saved.sort((a, b) => (b.netBenefitAllCohorts || 0) - (a.netBenefitAllCohorts || 0));
  return saved.slice(0, Math.min(limit, saved.length));
}

function getExportMode() {
  const brief = document.getElementById("export-mode-brief");
  const standard = document.getElementById("export-mode-standard");
  if (brief && brief.checked) return "brief";
  if (standard && standard.checked) return "standard";
  return "standard";
}

function getExportNotesFromUI() {
  const enablersEl = document.getElementById("export-enablers");
  const risksEl = document.getElementById("export-risks");
  return {
    enablers: enablersEl ? (enablersEl.value || "").trim() : "",
    risks: risksEl ? (risksEl.value || "").trim() : ""
  };
}



function renderBriefBulletsPreview() {
  const en = document.getElementById("export-enablers");
  const rk = document.getElementById("export-risks");
  const enPrev = document.getElementById("brief-enablers-preview");
  const rkPrev = document.getElementById("brief-risks-preview");
  if (!enPrev || !rkPrev) return;

  const enBullets = parseBulletsText(en ? en.value : "");
  const rkBullets = parseBulletsText(rk ? rk.value : "");

  enPrev.innerHTML = enBullets.length ? `<ul>${enBullets.map((b) => `<li>${escapeHtml(b)}</li>`).join("")}</ul>` : "<span class=\"muted\">(No enablers entered.)</span>";
  rkPrev.innerHTML = rkBullets.length ? `<ul>${rkBullets.map((b) => `<li>${escapeHtml(b)}</li>`).join("")}</ul>` : "<span class=\"muted\">(No risks entered.)</span>";
}

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}



function scenarioTableMarkdownRows(items) {
  const header =
    "| Scenario | Endorsement | Perceived programme value (all cohorts) | Economic cost (all cohorts) | Epidemiological benefit (all cohorts) | Net benefit | BCR | Feasibility |\n" +
    "|---|---:|---:|---:|---:|---:|---:|---|\n";
  const body = items
    .map((s) => {
      const feas = s.capacity ? s.capacity.status : computeCapacity(s.config).status;
      return `| ${safeText(getScenarioDisplayName(s))} | ${formatNumber(s.endorseRate, 1)}% | ${formatCurrencyDisplay(
        s.wtpAllCohorts,
        0
      )} | ${formatCurrencyDisplay(s.totalEconomicCostAllCohorts ?? s.natTotalCost, 0)} | ${formatCurrencyDisplay(s.totalEpiBenefitsAllCohorts ?? s.epiBenefitAllCohorts, 0)} | ${formatCurrencyDisplay(
        s.netBenefitAllCohorts,
        0
      )} | ${formatNumber((s.natBcr !== null && isFinite(s.natBcr)) ? s.natBcr : safeBcr(s), 2)} | ${safeText(feas)} |`;
    })
    .join("\n");
  return header + body;
}

function escapeHtmlSimple(x) {
  return safeText(x)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function updateValidationWarnings(config) {
  const container = document.getElementById("config-warnings");
  if (!container) return;

  const warnings = [];

  const totalTrainees = (Number(config.cohorts || 0) * Number(config.traineesPerCohort || 0)) || 0;
  if (totalTrainees >= 2000) {
    warnings.push(
      `Check realism: cohorts × trainees per cohort = ${formatNumber(totalTrainees, 0)}. Confirm staffing, sites, and scheduling assumptions.`
    );
  }

  const sliderEl = document.getElementById("cost-slider");
  if (sliderEl) {
    const min = Number(sliderEl.min);
    const max = Number(sliderEl.max);
    const v = Number(sliderEl.value);
    if (!isNaN(min) && !isNaN(max) && max > min) {
      const p = (v - min) / (max - min);
      if (p <= 0.05 || p >= 0.95) {
        warnings.push("Cost input is near the slider extreme. Sense-check whether the cost per trainee-month is realistic.");
      }
    }
  }

  if (warnings.length === 0) {
    container.innerHTML = "";
    container.style.display = "none";
    return;
  }

  container.style.display = "";
  container.innerHTML = warnings.map((w) => `<p class="warn">⚠ ${escapeHtmlSimple(w)}</p>`).join("");
}




function computeScenario(config) {
  const pref = computeEndorsementAndWTP(config);
  const costs = computeCosts(config);
  const epi = computeEpidemiological(config, pref.endorseRate);
  const capacity = computeCapacity(config);

  const wtpPerTraineePerMonth = pref.wtpPerTraineePerMonth;

  const wtpPerCohort = wtpPerTraineePerMonth * epi.months * config.traineesPerCohort;
  const wtpAllCohorts = wtpPerCohort * config.cohorts;

  const epiBenefitPerCohort = epi.epiBenefitPerCohort;
  const epiBenefitAllCohorts = epiBenefitPerCohort * config.cohorts;

  const netBenefitPerCohort = epiBenefitPerCohort - costs.totalEconomicCostPerCohort;
  const netBenefitAllCohorts = epiBenefitAllCohorts - costs.totalEconomicCostPerCohort * config.cohorts;

  const bcrPerCohort =
    costs.totalEconomicCostPerCohort > 0 ? epiBenefitPerCohort / costs.totalEconomicCostPerCohort : null;

  const natTotalCost = costs.totalEconomicCostPerCohort * config.cohorts;
  const natBcr = natTotalCost > 0 ? epiBenefitAllCohorts / natTotalCost : null;

  // Retain existing decomposition placeholder; unchanged computation
  const wtpOutbreakComponent = wtpAllCohorts * 0.3;

  return {
    id: Date.now().toString(),
    timestamp: new Date().toISOString(),
    shortlisted: false,
    config,
    preferenceModel: config.preferenceModel,
    endorseRate: pref.endorseRate,
    optOutRate: pref.optOutRate,

    // "Perceived programme value" (previously WTP) - calculations unchanged
    wtpPerTraineePerMonth,
    wtpPerCohort,
    wtpAllCohorts,
    wtpOutbreakComponent,

    costs,
    epi,
    capacity,

    epiBenefitPerCohort,
    epiBenefitAllCohorts,
    netBenefitPerCohort,
    netBenefitAllCohorts,
    bcrPerCohort,
    natTotalCost,
    natBcr,
    graduatesPerCohort: epi.graduatesPerCohort,
    graduatesAllCohorts: epi.graduatesAllCohorts,
    outbreaksPerYearPerCohort: epi.outbreaksPerYearPerCohort,
    outbreaksPerYearNational: epi.outbreaksPerYearNational,
    discountRate: epi.discountRate,
    planningYears: epi.planningYears
  };
}


/* ===========================
   Charts
   =========================== */

function ensureChart(ctxId, type, data, options) {
  if (!window.Chart) return null;
  const ctx = document.getElementById(ctxId)?.getContext("2d");
  if (!ctx) return null;
  return new Chart(ctx, { type, data, options });
}

function updateUptakeChart(scenario) {
  const ctxId = "chart-uptake";
  const existing = appState.charts.uptake;
  const data = {
    labels: ["Endorse FETP option", "Choose opt out"],
    datasets: [{ label: "Share of stakeholders", data: [scenario.endorseRate, scenario.optOutRate] }]
  };
  const options = {
    responsive: true,
    plugins: { legend: { display: false } },
    scales: { y: { beginAtZero: true, max: 100 } }
  };
  if (existing) {
    existing.data = data;
    existing.options = options;
    existing.update();
  } else {
    appState.charts.uptake = ensureChart(ctxId, "bar", data, options);
  }
}

function updateBcrChart(scenario) {
  const ctxId = "chart-bcr";
  const existing = appState.charts.bcr;
  const data = {
    labels: ["Indicative outbreak cost saving", "Economic cost"],
    datasets: [
      {
        label: "Per cohort (INR)",
        data: [scenario.epiBenefitPerCohort, scenario.costs.totalEconomicCostPerCohort]
      }
    ]
  };
  const options = {
    responsive: true,
    plugins: { legend: { display: false } },
    scales: {
      y: {
        beginAtZero: true,
        ticks: { callback: (value) => formatNumber(value, 0) }
      }
    }
  };
  if (existing) {
    existing.data = data;
    existing.options = options;
    existing.update();
  } else {
    appState.charts.bcr = ensureChart(ctxId, "bar", data, options);
  }
}

function updateEpiChart(scenario) {
  const ctxId = "chart-epi";
  const existing = appState.charts.epi;
  const data = {
    labels: ["Graduates (all cohorts)", "Outbreak responses per year"],
    datasets: [{ label: "Epidemiological outputs", data: [scenario.graduatesAllCohorts, scenario.outbreaksPerYearNational] }]
  };
  const options = { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } };
  if (existing) {
    existing.data = data;
    existing.options = options;
    existing.update();
  } else {
    appState.charts.epi = ensureChart(ctxId, "bar", data, options);
  }
}

function updateNatCostBenefitChart(scenario) {
  const ctxId = "chart-nat-cost-benefit";
  const existing = appState.charts.natCostBenefit;
  const totalBenefit = scenario.epiBenefitAllCohorts;
  const data = {
    labels: ["Total economic cost (all cohorts)", "Total outbreak cost saving (all cohorts)"],
    datasets: [{ label: "National totals (INR)", data: [scenario.natTotalCost, totalBenefit] }]
  };
  const options = {
    responsive: true,
    plugins: { legend: { display: false } },
    scales: {
      y: { beginAtZero: true, ticks: { callback: (value) => formatNumber(value, 0) } }
    }
  };
  if (existing) {
    existing.data = data;
    existing.options = options;
    existing.update();
  } else {
    appState.charts.natCostBenefit = ensureChart(ctxId, "bar", data, options);
  }
}

function updateNatEpiChart(scenario) {
  const ctxId = "chart-nat-epi";
  const existing = appState.charts.natEpi;
  const data = {
    labels: ["Total graduates", "Outbreak responses per year"],
    datasets: [{ label: "National epidemiological outputs", data: [scenario.graduatesAllCohorts, scenario.outbreaksPerYearNational] }]
  };
  const options = { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } };
  if (existing) {
    existing.data = data;
    existing.options = options;
    existing.update();
  } else {
    appState.charts.natEpi = ensureChart(ctxId, "bar", data, options);
  }
}

/* ===========================
   UI updates
   =========================== */

function updateCostSliderLabel() {
  const slider = document.getElementById("cost-slider");
  const display = document.getElementById("cost-display");
  if (!slider || !display) return;
  const val = Number(slider.value);
  display.textContent = formatCurrencyDisplay(val, 0);
}

function updateCurrencyToggle() {
  const label = document.getElementById("currency-label");
  const buttons = Array.from(document.querySelectorAll(".pill-toggle"));
  buttons.forEach((btn) => {
    const c = btn.getAttribute("data-currency");
    if (c === appState.currency) btn.classList.add("active");
    else btn.classList.remove("active");
  });
  if (label) label.textContent = appState.currency;
  if (appState.currentScenario) refreshAllOutputs(appState.currentScenario);
}

/* ===========================
   Settings tab (UI contract)
   =========================== */

function readSettingsFormValues() {
  const form = document.getElementById("settingsForm");
  if (!form) return null;

  const inputs = Array.from(form.querySelectorAll("input, select, textarea"));
  const values = {};

  inputs.forEach((el) => {
    if (!el || el.disabled) return;
    const key = el.name || el.id;
    if (!key) return;

    let value = null;
    if (el.type === "checkbox") value = !!el.checked;
    else value = el.value;

    values[key] = value;
  });

  return values;
}

function applySettingsValuesToState(values) {
  if (!values) return;

  const general = appState.epiSettings.general;
  const tiers = appState.epiSettings.tiers;

  const applyToAllTiers = (fn) => {
    ["frontline", "intermediate", "advanced"].forEach((tier) => fn(tiers[tier], tier));
  };

  Object.keys(values).forEach((k) => {
    const raw = values[k];
    const lower = String(k).toLowerCase();
    const num = Number(raw);

    if (lower.includes("planning") && lower.includes("horizon")) {
      const v = Number(raw);
      if (isFinite(v) && v > 0) general.planningHorizonYears = Math.round(v);
      return;
    }

    if (lower.includes("discount")) {
      const v = Number(raw);
      if (isFinite(v) && v >= 0) {
        let r = v;
        if (r > 1) r = r / 100;
        general.epiDiscountRate = clamp(r, 0, 1);
      }
      return;
    }

    if ((lower.includes("usd") || lower.includes("exchange") || lower.includes("inr_to_usd") || lower.includes("inrtousd"))) {
      const v = Number(raw);
      if (isFinite(v) && v > 0) {
        general.inrToUsdRate = v;
        appState.usdRate = v;
      }
      return;
    }

    if (lower.includes("value") && lower.includes("outbreak")) {
      const vInr = parseSensitivityValueToINR(raw);
      if (vInr && isFinite(vInr) && vInr > 0) {
        applyToAllTiers((t) => {
          t.valuePerOutbreak = vInr;
        });
      } else if (isFinite(num) && num > 0) {
        let v = num;
        if (v < 1000) v = v * 1e6;
        applyToAllTiers((t) => {
          t.valuePerOutbreak = v;
        });
      }
      return;
    }

    if (lower.includes("completion") && lower.includes("rate")) {
      if (isFinite(num) && num >= 0) {
        let cr = num;
        if (cr > 1) cr = cr / 100;
        cr = clamp(cr, 0, 1);
        applyToAllTiers((t) => {
          t.completionRate = cr;
        });
      }
      return;
    }

    if (lower.includes("outbreaks") && lower.includes("graduate")) {
      if (isFinite(num) && num >= 0) {
        applyToAllTiers((t) => {
          t.outbreaksPerGraduatePerYear = num;
        });
      }
      return;
    }

    if (lower.includes("value") && lower.includes("graduate")) {
      if (isFinite(num) && num >= 0) {
        applyToAllTiers((t) => {
          t.valuePerGraduate = num;
        });
      }
      return;
    }
  });

  appState.settings.lastAppliedValues = values;

  syncOutbreakValueDropdownsFromState();
}

function buildHumanReadableSettingsSummary() {
      const g = appState.epiSettings.general;
      const t = appState.epiSettings.tiers.frontline;

      const parts = [];
      parts.push(`Planning horizon ${formatNumber(g.planningHorizonYears, 0)} years`);
      parts.push(`Discount rate ${formatNumber(g.epiDiscountRate * 100, 1)} percent`);
      parts.push(`INR per USD ${formatNumber(g.inrToUsdRate, 2)}`);
      parts.push(`Value per outbreak ₹${formatNumber(t.valuePerOutbreak, 0)}`);
      parts.push(`Completion rate ${formatNumber(t.completionRate * 100, 1)} percent`);
      parts.push(`Outbreaks per graduate per year ${formatNumber(t.outbreaksPerGraduatePerYear, 2)}`);

      const mentorCostEl = document.getElementById("mentor-support-cost-per-cohort");
      const mentorsAvailEl = document.getElementById("available-mentors-national");
      const sitesAvailEl = document.getElementById("available-training-sites");
      const maxCohortsSiteEl = document.getElementById("max-cohorts-per-site");
      const crossSectorEl = document.getElementById("cross-sector-multiplier");

      const capPieces = [];
      if (mentorCostEl) {
        const v = Number(mentorCostEl.value.replace(/,/g, "")) || 0;
        capPieces.push(`Mentor support cost base per cohort ₹${formatNumber(v, 0)}`);
      }
      if (mentorsAvailEl) {
        const v = Number(mentorsAvailEl.value.replace(/,/g, "")) || 0;
        capPieces.push(`Available mentors nationally ${formatNumber(v, 0)}`);
      }
      if (sitesAvailEl) {
        const v = Number(sitesAvailEl.value.replace(/,/g, "")) || 0;
        capPieces.push(`Available training sites / hubs ${formatNumber(v, 0)}`);
      }
      if (maxCohortsSiteEl) {
        const v = Number(maxCohortsSiteEl.value.replace(/,/g, "")) || 0;
        capPieces.push(`Max cohorts per site per year ${formatNumber(v, 0)}`);
      }
      if (crossSectorEl) {
        const v = Number(crossSectorEl.value.replace(/,/g, "")) || 0;
        capPieces.push(`Cross-sector benefit multiplier ${formatNumber(v, 2)}×`);
      }
      if (capPieces.length) {
        parts.push("Capacity and costs: " + capPieces.join("; "));
      }

      return parts.join("; ");
    }


function appendSettingsLogEntry(text) {
  const time = new Date().toLocaleString();

  const targets = [];
  const contractLog = document.getElementById("settingsLog");
  const sessionLog = document.getElementById("settings-log");
  const advLog = document.getElementById("adv-settings-log");

  if (contractLog) targets.push(contractLog);
  if (sessionLog && sessionLog !== contractLog) targets.push(sessionLog);
  if (advLog && advLog !== sessionLog && advLog !== contractLog) targets.push(advLog);

  if (!targets.length) return;

  targets.forEach((log) => {
    const entry = document.createElement("div");
    entry.className = "settings-log-entry";
    entry.textContent = `[${time}] ${text}`;
    log.appendChild(entry);
    log.scrollTop = log.scrollHeight;
  });
}

function initApplySettingsButton() {
  const btn = document.getElementById("applySettingsBtn");
  if (!btn) return;

  btn.disabled = false;
  btn.removeAttribute("aria-disabled");

  btn.addEventListener("click", () => {
    const values = readSettingsFormValues();
    applySettingsValuesToState(values);

    if (appState.currentScenario) {
      const c = { ...appState.currentScenario.config };
      c.planningHorizonYears = appState.epiSettings.general.planningHorizonYears;
      const salaryInputs = readSalaryInputsFromDom();
      c.salaryMonthlyFaculty = salaryInputs.facultyMonthly;
      c.salaryMonthlyCoordinator = salaryInputs.coordinatorMonthly;
      c.salaryMonthlyParticipant = salaryInputs.participantMonthly;
      const newScenario = computeScenario(c);
      appState.currentScenario = newScenario;
      refreshAllOutputs(newScenario);
    }

    const summary = buildHumanReadableSettingsSummary();
    appendSettingsLogEntry(`Settings applied. ${summary}.`);
    showToast("Settings applied.", "success");

    syncOutbreakValueDropdownsFromState();
  });
}

/* ===========================
   Results and national tabs updates
   =========================== */

function updateConfigSummary(scenario) {
  const container = document.getElementById("config-summary");
  if (!container) return;

  const c = scenario.config;
  container.innerHTML = "";

  const rows = [
    {
      label: "Programme tier",
      value: c.tier === "frontline" ? "Frontline" : c.tier === "intermediate" ? "Intermediate" : "Advanced"
    },
    {
      label: "Career incentive",
      value:
        c.career === "certificate"
          ? "Government and partner certificate"
          : c.career === "uniqual"
          ? "University qualification"
          : "Government career pathway"
    },
    {
      label: "Mentorship intensity",
      value: c.mentorship === "low" ? "Low" : c.mentorship === "medium" ? "Medium" : "High"
    },
    {
      label: "Delivery mode",
      value: c.delivery === "blended" ? "Blended" : c.delivery === "inperson" ? "Fully in person" : "Fully online"
    },
    { label: "Response time", value: "Detect and respond within 7 days" },
    { label: "Cost per trainee per month", value: formatCurrencyDisplay(c.costPerTraineePerMonth, 0) },
    { label: "Trainees per cohort", value: formatNumber(c.traineesPerCohort, 0) },
    { label: "Number of cohorts", value: formatNumber(c.cohorts, 0) },
    {
      label: "Planning horizon (years)",
      value: formatNumber(c.planningHorizonYears || appState.epiSettings.general.planningHorizonYears, 0)
    },
    { label: "Opportunity cost", value: c.opportunityCostIncluded ? "Included in economic cost" : "Not included" }
  ];

  rows.forEach((row) => {
    const div = document.createElement("div");
    div.className = "config-summary-row";
    div.innerHTML = `
      <span class="config-summary-label">${row.label}</span>
      <span class="config-summary-value">${row.value}</span>
    `;
    container.appendChild(div);
  });

  const endorsementEl = document.getElementById("config-endorsement-value");
  if (endorsementEl) endorsementEl.textContent = formatNumber(scenario.endorseRate, 1) + "%";

  const statusTag = document.getElementById("headline-status-tag");
  if (statusTag) {
    statusTag.textContent = "";
    statusTag.classList.remove("status-neutral", "status-good", "status-warning", "status-poor");

    let statusClass = "status-neutral";
    let statusText = "Scenario assessed";

    if (scenario.endorseRate >= 70 && scenario.bcrPerCohort !== null && scenario.bcrPerCohort >= 1.2) {
      statusClass = "status-good";
      statusText = "Strong configuration";
    } else if (scenario.endorseRate >= 50 && scenario.bcrPerCohort !== null && scenario.bcrPerCohort >= 1.0) {
      statusClass = "status-warning";
      statusText = "Promising configuration (needs discussion)";
    } else {
      statusClass = "status-poor";
      statusText = "Challenging configuration (lower support and perceived value below cost)";
    }

    statusTag.classList.add(statusClass);
    statusTag.textContent = statusText;
  }

  const headlineText = document.getElementById("headline-recommendation");
  if (headlineText) {
    const endorse = formatNumber(scenario.endorseRate, 1);
    const cost = formatCurrencyDisplay(scenario.costs.totalEconomicCostPerCohort, 0);
    const bcr = scenario.bcrPerCohort !== null ? formatNumber(scenario.bcrPerCohort, 2) : "-";
    headlineText.textContent =
      `The mixed logit preference model points to an endorsement rate of about ${endorse} percent, an economic cost of ${cost} per cohort and an indicative outbreak cost saving to cost ratio near ${bcr}. These values give a concise starting point for discussions with ministries and partners.`;
  }

  const briefingEl = document.getElementById("headline-briefing-text");
  if (briefingEl) {
    const natCost = formatCurrencyDisplay(scenario.natTotalCost, 0);
    const natBenefit = formatCurrencyDisplay(scenario.epiBenefitAllCohorts, 0);
    const natBcr = scenario.natBcr !== null ? formatNumber(scenario.natBcr, 2) : "-";
    const statusQuoLine =
      "Status quo: benchmark needs imply around 6,500 Intermediate/Advanced field epidemiologists; STEPS works with current stock around 3,300, implying a gap near 3,200 Intermediate/Advanced-equivalent graduates by 2030. STEPS allows testing alternative configurations to close this gap.";

    briefingEl.textContent =
      `${statusQuoLine}\n\nWith this configuration, about ${formatNumber(scenario.endorseRate, 1)} percent of stakeholders are expected to endorse the investment. Running ${formatNumber(
        scenario.config.cohorts,
        0
      )} cohorts of ${formatNumber(scenario.config.traineesPerCohort, 0)} trainees leads to a total economic cost of roughly ${natCost} over the planning horizon and an indicative outbreak related economic cost saving of roughly ${natBenefit}. The national benefit cost ratio is around ${natBcr}, based on the outbreak value and epidemiological assumptions set in the settings and methods.`;
  }
}

function updateResultsTab(scenario) {
  const endorseEl = document.getElementById("endorsement-rate");
  const optOutEl = document.getElementById("optout-rate");
  const wtpPerTraineeEl = document.getElementById("wtp-per-trainee");
  const wtpTotalCohortEl = document.getElementById("wtp-total-cohort");
  const progCostEl = document.getElementById("prog-cost-per-cohort");
  const totalCostEl = document.getElementById("total-cost");
  const netBenefitEl = document.getElementById("net-benefit");
  const bcrEl = document.getElementById("bcr");
  const gradsEl = document.getElementById("epi-graduates");
  const outbreaksEl = document.getElementById("epi-outbreaks");
  const epiBenefitEl = document.getElementById("epi-benefit");

  // NEW elements (defensive)
  const mentorCostEl = document.getElementById("mentor-cost-per-cohort");
  const directCostEl = document.getElementById("direct-cost-per-cohort");

  const existingOppCostEl = document.getElementById("existing-opp-cost-per-cohort");
  const salaryOppCostEl = document.getElementById("salary-opp-cost-per-cohort");

  // Capacity panel elements (defensive)
  const capMentorsPerCohortEl = document.getElementById("capacity-mentors-per-cohort");
  const capTotalMentorsEl = document.getElementById("capacity-total-mentors");
  const capAvailableMentorsEl = document.getElementById("capacity-available-mentors");
  const capShortfallEl = document.getElementById("capacity-mentor-shortfall");
  const capStatusEl = document.getElementById("capacity-status");
  const capNoteEl = document.getElementById("capacity-note");
  const capSiteCapacityEl = document.getElementById("capacity-site-capacity");
  const capSiteGapEl = document.getElementById("capacity-site-gap");
  const capSitesRow = document.getElementById("capacity-sites-row");

  if (endorseEl) endorseEl.textContent = formatNumber(scenario.endorseRate, 1) + "%";
  if (optOutEl) optOutEl.textContent = formatNumber(scenario.optOutRate, 1) + "%";

  // Relabelled in UI as perceived programme value; IDs retained for backwards compatibility
  if (wtpPerTraineeEl) wtpPerTraineeEl.textContent = formatCurrencyDisplay(scenario.wtpPerTraineePerMonth, 0);
  if (wtpTotalCohortEl) wtpTotalCohortEl.textContent = formatCurrencyDisplay(scenario.wtpPerCohort, 0);

  if (progCostEl) progCostEl.textContent = formatCurrencyDisplay(scenario.costs.programmeCostPerCohort, 0);
  if (mentorCostEl) mentorCostEl.textContent = formatCurrencyDisplay(scenario.costs.mentorCostPerCohort, 0);
  if (directCostEl) directCostEl.textContent = formatCurrencyDisplay(scenario.costs.directCostPerCohort, 0);

  if (existingOppCostEl) existingOppCostEl.textContent = formatCurrencyDisplay(scenario.costs.opportunityCostPerCohort, 0);
  if (salaryOppCostEl) salaryOppCostEl.textContent = formatCurrencyDisplay(scenario.costs.salaryBasedOpportunityCostPerCohort, 0);

  // Existing "total-cost" remains economic cost per cohort
  if (totalCostEl) totalCostEl.textContent = formatCurrencyDisplay(scenario.costs.totalEconomicCostPerCohort, 0);

  if (netBenefitEl) netBenefitEl.textContent = formatCurrencyDisplay(scenario.netBenefitPerCohort, 0);
  if (bcrEl) bcrEl.textContent = scenario.bcrPerCohort !== null ? formatNumber(scenario.bcrPerCohort, 2) : "-";

  if (gradsEl) gradsEl.textContent = formatNumber(scenario.graduatesAllCohorts, 0);
  if (outbreaksEl) outbreaksEl.textContent = formatNumber(scenario.outbreaksPerYearNational, 1);
  if (epiBenefitEl) epiBenefitEl.textContent = formatCurrencyDisplay(scenario.epiBenefitPerCohort, 0);

  // Capacity/feasibility
  if (scenario.capacity) {
    if (capMentorsPerCohortEl) capMentorsPerCohortEl.textContent = formatNumber(scenario.capacity.mentorsPerCohort, 0);
    if (capTotalMentorsEl) capTotalMentorsEl.textContent = formatNumber(scenario.capacity.totalMentorsRequired, 0);
        const capTotalMentorsNationalEl = document.getElementById("capacity-total-mentors-required");
        if (capTotalMentorsNationalEl) capTotalMentorsNationalEl.textContent = formatNumber(scenario.capacity.totalMentorsRequired, 0);
    if (capAvailableMentorsEl) capAvailableMentorsEl.textContent = formatNumber(scenario.capacity.availableMentors, 0);
    if (capShortfallEl) capShortfallEl.textContent = formatNumber(scenario.capacity.mentorShortfall, 0);
    if (capStatusEl) capStatusEl.textContent = scenario.capacity.status || "-";

    if (capNoteEl) {
      if (scenario.capacity.mentorShortfall > 0) {
        capNoteEl.textContent = `Mentor gap of ${formatNumber(scenario.capacity.mentorShortfall, 0)} to meet the selected cohort and mentorship intensity.`;
      } else {
        capNoteEl.textContent = "Mentor capacity appears sufficient for the selected cohorts and mentorship intensity.";
      }
    }

    const showSites = scenario.capacity.siteCapacity !== null && scenario.capacity.siteCapacity !== undefined;
    if (capSitesRow) capSitesRow.style.display = showSites ? "" : "none";
    if (showSites) {
      if (capSiteCapacityEl) capSiteCapacityEl.textContent = formatNumber(scenario.capacity.siteCapacity, 1);
      if (capSiteGapEl) capSiteGapEl.textContent = formatNumber(scenario.capacity.siteGap, 1);
    }
  }
}


function updateCostingTab(scenario) {
  const select = document.getElementById("cost-source");
  if (select && select.options.length === 0) {
    ["frontline", "intermediate", "advanced"].forEach((tier) => {
      const templates = COST_TEMPLATES[tier];
      if (templates && templates.combined) {
        const opt = document.createElement("option");
        opt.value = templates.combined.id;
        opt.textContent = templates.combined.label;
        select.appendChild(opt);
      }
    });
  }

  if (select) {
    const templates = COST_TEMPLATES[scenario.config.tier];
    if (templates && templates.combined) select.value = templates.combined.id;
  }

  const summaryBox = document.getElementById("cost-breakdown-summary");
  const tbody = document.getElementById("cost-components-list");
  if (!summaryBox || !tbody) return;

  tbody.innerHTML = "";
  summaryBox.innerHTML = "";

  const costInfo = scenario.costs;
  const template = costInfo.template;
  const directCost = costInfo.programmeCostPerCohort;
  const oppCost = costInfo.opportunityCostPerCohort;
  const salaryOppCost = costInfo.salaryBasedOpportunityCostPerCohort || 0;
  const salaryOppCostRaw = costInfo.salaryBasedOpportunityCostPerCohortRaw || 0;
  const econCost = costInfo.totalEconomicCostPerCohort;

  const cardsData = [
    { label: "Programme cost per cohort", value: formatCurrencyDisplay(directCost, 0) },
    { label: "Mentor support cost per cohort", value: formatCurrencyDisplay(costInfo.mentorCostPerCohort || 0, 0) },
    { label: "Direct cost per cohort", value: formatCurrencyDisplay(costInfo.directCostPerCohort || 0, 0) },
    { label: "Existing opportunity cost per cohort", value: formatCurrencyDisplay(oppCost, 0) },
    { label: "Salary-based opportunity cost per cohort (additional)", value: formatCurrencyDisplay(salaryOppCost, 0) },
    { label: "Total economic cost per cohort", value: formatCurrencyDisplay(econCost, 0) },
    {
      label: "Share of indirect cost (existing + salary-based)",
      value: econCost > 0 ? formatNumber(((oppCost + salaryOppCost) / econCost) * 100, 1) + "%" : "-"
    },
    {
      label: "Salary-based excluded amount (if switch off)",
      value: (!scenario.config.opportunityCostIncluded && salaryOppCostRaw > 0) ? formatCurrencyDisplay(salaryOppCostRaw, 0) : "Not provided"
    }
  ];

  cardsData.forEach((c) => {
    const div = document.createElement("div");
    div.className = "cost-summary-card";
    div.innerHTML = `
      <div class="cost-summary-label">${c.label}</div>
      <div class="cost-summary-value">${c.value}</div>
    `;
    summaryBox.appendChild(div);
  });

  if (!template) return;

  const months = TIER_MONTHS[scenario.config.tier] || 12;
  const trainees = scenario.config.traineesPerCohort;
  const directForComponents = directCost;

  template.components.forEach((comp) => {
    const amount = directForComponents * comp.directShare;
    const perTraineePerMonth = trainees > 0 && months > 0 ? amount / (trainees * months) : 0;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${comp.label}</td>
      <td class="numeric-cell">${formatNumber(comp.directShare * 100, 1)}%</td>
      <td class="numeric-cell">${formatCurrencyDisplay(amount, 0)}</td>
      <td class="numeric-cell">${formatCurrencyDisplay(perTraineePerMonth, 0)}</td>
      <td>Included in combined template for this tier.</td>
    `;
    tbody.appendChild(tr);
  });


  // NEW: Mentor support cost as explicit component (per cohort)
  if (tbody && scenario.costs) {
    const mentorPerCohort = Number(scenario.costs.mentorCostPerCohort || 0);
    if (mentorPerCohort > 0) {
      const trainees = Number(scenario.config.traineesPerCohort || 0);
      const months = TIER_MONTHS[scenario.config.tier] || 12;
      const perTraineePerMonth = trainees > 0 && months > 0 ? mentorPerCohort / (trainees * months) : 0;
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>Mentor support (explicit)</td>
        <td class="numeric-cell">Not provided</td>
        <td class="numeric-cell">${formatCurrencyDisplay(mentorPerCohort, 0)}</td>
        <td class="numeric-cell">${formatCurrencyDisplay(perTraineePerMonth, 0)}</td>
        <td>Base mentor support cost per cohort scaled by mentorship intensity multiplier.</td>
      `;
      tbody.appendChild(tr);
    }
  }


  // Existing opportunity cost as explicit economic component (per cohort)
  if (tbody && scenario.costs) {
    const trainees = Number(scenario.config.traineesPerCohort || 0);
    const months = TIER_MONTHS[scenario.config.tier] || 12;

    const section = document.createElement("tr");
    section.className = "section-row";
    section.innerHTML = `<td colspan="5">Economic (non-financial) cost components</td>`;
    tbody.appendChild(section);

    const opp = Number(scenario.costs.opportunityCostPerCohort || 0);
    const oppPerTraineePerMonth = trainees > 0 && months > 0 ? opp / (trainees * months) : 0;
    const trOpp = document.createElement("tr");
    trOpp.innerHTML = `
      <td>Existing opportunity cost (template-based)</td>
      <td class="numeric-cell">Not provided</td>
      <td class="numeric-cell">${formatCurrencyDisplay(opp, 0)}</td>
      <td class="numeric-cell">${formatCurrencyDisplay(oppPerTraineePerMonth, 0)}</td>
      <td>Existing STEPS opportunity cost methodology (unchanged).</td>
    `;
    tbody.appendChild(trOpp);

    const salaryTotalIncluded = Number(scenario.costs.salaryBasedOpportunityCostPerCohort || 0);
    const salaryRaw = Number(scenario.costs.salaryBasedOpportunityCostPerCohortRaw || 0);
    const salaryMeta = scenario.costs.salaryBasedOpportunityCostMeta || {};
    const salaryByRole = scenario.costs.salaryBasedOpportunityCostByRole || {};
    const salaryTotalPerTraineePerMonth = trainees > 0 && months > 0 ? salaryTotalIncluded / (trainees * months) : 0;

    const trSalary = document.createElement("tr");
    trSalary.innerHTML = `
      <td>Salary-based opportunity cost (additional) Not provided total</td>
      <td class="numeric-cell">Not provided</td>
      <td class="numeric-cell">${formatCurrencyDisplay(salaryTotalIncluded, 0)}</td>
      <td class="numeric-cell">${formatCurrencyDisplay(salaryTotalPerTraineePerMonth, 0)}</td>
      <td>Based on monthly salary inputs and contact day assumptions (see Settings and appendix). ${scenario.config.opportunityCostIncluded ? "" : "Computed but excluded from totals because the opportunity cost switch is off."}</td>
    `;
    tbody.appendChild(trSalary);

    // Role breakdown rows (salary-based)
    const roleRows = [
      ["Faculty (proxy: mentors per cohort)", Number(salaryByRole.faculty || 0)],
      ["Coordinator (assumed 1 per cohort)", Number(salaryByRole.coordinator || 0)],
      ["Participants (trainees per cohort)", Number(salaryByRole.participant || 0)]
    ];
    roleRows.forEach(([label, val]) => {
      const perTrainee = trainees > 0 && months > 0 ? val / (trainees * months) : 0;
      const tr = document.createElement("tr");
      tr.className = "subrow";
      tr.innerHTML = `
        <td>&nbsp;&nbsp;${label}</td>
        <td class="numeric-cell">Not provided</td>
        <td class="numeric-cell">${formatCurrencyDisplay(val, 0)}</td>
        <td class="numeric-cell">${formatCurrencyDisplay(perTrainee, 0)}</td>
        <td>Salary-based method (additional). Delivery: ${safeText(salaryMeta.delivery || scenario.config.delivery || "")}; Contact days: ${formatNumber(salaryMeta.contactDays ?? "", 0)}; Duration: ${formatNumber(salaryMeta.months ?? months, 0)} months.</td>
      `;
      tbody.appendChild(tr);
    });
  }

}


function updateNationalSimulationTab(scenario) {
  // National Simulation is portfolio-based: baseline tiers + configured tier override.
  const horizon = getPlanningHorizonFromInputs();
  const natScenario = (scenario && scenario.config && scenario.config.isPortfolio)
    ? scenario
    : computeScenarioAggregateFromCurrentConfig(horizon);

  const totCostEl = document.getElementById("nat-total-cost");
  const totBenefitEl = document.getElementById("nat-total-benefit");
  const netBenefitEl = document.getElementById("nat-net-benefit");
  const natBcrEl = document.getElementById("nat-bcr");
  const natGraduatesEl = document.getElementById("nat-graduates");
  const natOutbreaksEl = document.getElementById("nat-outbreaks");
  const natTotalWtpEl = document.getElementById("nat-total-wtp");
  const textEl = document.getElementById("natsim-summary-text");

  const natCost = natScenario.natTotalCost;
  const natBenefit = natScenario.epiBenefitAllCohorts;
  const natNet = natScenario.netBenefitAllCohorts;
  const natBcr = natScenario.natBcr !== null ? natScenario.natBcr : null;
  const natTotalWtp = natScenario.wtpAllCohorts;

  if (totCostEl) totCostEl.textContent = formatCurrencyDisplay(natCost, 0);
  if (totBenefitEl) totBenefitEl.textContent = formatCurrencyDisplay(natBenefit, 0);
  if (netBenefitEl) netBenefitEl.textContent = formatCurrencyDisplay(natNet, 0);
  if (natBcrEl) natBcrEl.textContent = natBcr !== null ? formatNumber(natBcr, 2) : "-";
  if (natGraduatesEl) natGraduatesEl.textContent = formatNumber(natScenario.graduatesAllCohorts, 0);
  if (natOutbreaksEl) natOutbreaksEl.textContent = formatNumber(natScenario.outbreaksPerYearNational, 1);
  if (natTotalWtpEl) natTotalWtpEl.textContent = formatCurrencyDisplay(natTotalWtp, 0);

  if (textEl) {
    textEl.textContent =
      `At national level, this configuration (baseline tiers plus the currently configured ${safeText(natScenario._selectedTier || "")} tier) would produce about ${formatNumber(
        natScenario.graduatesAllCohorts,
        0
      )} graduates over the selected timeframe and support around ${formatNumber(
        natScenario.outbreaksPerYearNational,
        1
      )} outbreak responses per year once all cohorts are complete. The total economic cost across all cohorts is roughly ${formatCurrencyDisplay(
        natCost,
        0
      )}, while the indicative outbreak related economic cost saving is roughly ${formatCurrencyDisplay(
        natBenefit,
        0
      )}. This implies a national benefit cost ratio of about ${natBcr !== null ? formatNumber(natBcr, 2) : "-"} and a net outbreak related cost saving of ${formatCurrencyDisplay(
        natNet,
        0
      )}. Total willingness to pay across all cohorts is roughly ${formatCurrencyDisplay(natTotalWtp, 0)}.`;
  }

  updateNatCostBenefitChart(natScenario);
  updateNatEpiChart(natScenario);

  try { if (typeof renderNationalBaselineScenarioDelta === 'function') renderNationalBaselineScenarioDelta(null); } catch (e) { /* non-fatal */ }
}


/* ===========================
   Scenarios table and exports
   =========================== */

function refreshSavedScenariosTable() {
  const tbody = document.querySelector("#scenario-table tbody");
  if (!tbody) return;

  tbody.innerHTML = "";

  const headerEls = Array.from(document.querySelectorAll("#scenario-table thead th"));
  const headers = headerEls.map((h) => (h.textContent || "").trim().toLowerCase());

  function cellNumeric(val, decimals = 0) {
    return `<td class="numeric-cell">${formatNumber(val, decimals)}</td>`;
  }
  function cellCurrency(val, decimals = 0) {
    return `<td class="numeric-cell">${formatCurrencyDisplay(val, decimals)}</td>`;
  }

  appState.savedScenarios.forEach((s) => {
    const c = s.config;

    const isPortfolio = !!c?.isPortfolio || (String(c?.tier || "").toLowerCase() === "portfolio") || !!c?.portfolioTiers;
    let portfolioSummary = null;
    if (isPortfolio && c?.portfolioTiers) {
      const tiers = c.portfolioTiers;
      const keys = ["frontline","intermediate","advanced"];
      let cohortsSum = 0;
      let traineesSum = 0;
      let costWeightedByTrainees = 0;

      keys.forEach((k) => {
        const t = tiers[k];
        if (!t) return;
        const coh = Math.max(0, Number(t.cohorts || 0));
        const trn = Math.max(0, Number(t.traineesPerCohort || 0));
        const cost = Math.max(0, Number(t.costPerTraineePerMonth || 0));
        cohortsSum += coh;
        traineesSum += coh * trn;
        costWeightedByTrainees += (coh * trn) * cost;
      });

      portfolioSummary = {
        cohorts: cohortsSum,
        traineesPerCohortAvg: cohortsSum > 0 ? (traineesSum / cohortsSum) : 0,
        costPerTraineePerMonthAvg: traineesSum > 0 ? (costWeightedByTrainees / traineesSum) : 0
      };
    }
    const tierLabel = isPortfolio ? "Portfolio" : (c.tier === "advanced" ? "Advanced" : c.tier === "intermediate" ? "Intermediate" : c.tier === "frontline" ? "Frontline" : safeText(c.tier || ""));

    const careerLabel = isPortfolio ? "Mixed" : (c.career === "certificate" ? "Certificate" : c.career === "uniqual" ? "University qualification" : c.career === "career_path" ? "Career pathway" : safeText(c.career || ""));

    const mentorshipLabel = isPortfolio ? "Mixed" : (c.mentorship === "high" ? "High" : c.mentorship === "low" ? "Low" : c.mentorship === "medium" ? "Medium" : safeText(c.mentorship || ""));

    const deliveryLabel = isPortfolio ? "Mixed" : (c.delivery === "blended" ? "Blended" : c.delivery === "inperson" ? "Fully in person" : c.delivery === "online" ? "Fully online" : safeText(c.delivery || ""));

    const responseLabel = "Within 7 days";

    const cap = s.capacity || computeCapacity(c);

    const tagsHtml = `
      <td class="scenario-tags">
        <span class="tag-pill">${tierLabel}</span>
        <span class="tag-pill">${mentorshipLabel}</span>
        <span class="tag-pill">${deliveryLabel}</span>
      </td>
    `;

    const cellsByHeader = (h) => {
      if (h.includes("shortlist")) {
        return `<td><input type="checkbox" data-scenario-id="${s.id}" aria-label="Shortlist scenario" ${s.shortlisted ? "checked" : ""}></td>`;
      }
      if (h.includes("pin")) {
        return `<td><input type="checkbox" data-pin-id="${s.id}" aria-label="Pin scenario" ${s.pinned ? "checked" : ""}></td>`;
      }
      if (h === "name" || h.includes("scenario name")) {
        return `<td>${safeText(c.name)}</td>`;
      }
      if (h.includes("tag")) return tagsHtml;
      if (h.includes("tier")) return `<td>${tierLabel}</td>`;
      if (h.includes("career")) return `<td>${careerLabel}</td>`;
      if (h.includes("mentor") && h.includes("intensity")) return `<td>${mentorshipLabel}</td>`;
      if (h.includes("mentorship") && !h.includes("cost")) return `<td>${mentorshipLabel}</td>`;
      if (h.includes("delivery")) return `<td>${deliveryLabel}</td>`;
      if (h.includes("response")) return `<td>${responseLabel}</td>`;
      if (h === "cohorts" || h.includes("number of cohorts")) return cellNumeric(isPortfolio ? (portfolioSummary?.cohorts ?? 0) : c.cohorts, 0);
      if (h.includes("trainees") && h.includes("cohort")) return cellNumeric(isPortfolio ? (portfolioSummary?.traineesPerCohortAvg ?? 0) : c.traineesPerCohort, 0);
      if (h.includes("cost per trainee") && h.includes("month")) return cellCurrency(isPortfolio ? (portfolioSummary?.costPerTraineePerMonthAvg ?? 0) : c.costPerTraineePerMonth, 0);

      if (h.includes("mentor cost") && h.includes("per cohort")) return cellCurrency(s.costs?.mentorCostPerCohort ?? 0, 0);
      if (h.includes("total mentor")) return cellCurrency((s.costs?.mentorCostPerCohort ?? 0) * (c.cohorts || 0), 0);

      if (h.includes("direct cost") && (h.includes("total") || h.includes("all cohorts") || h.includes("national"))) {
        return cellCurrency(s.costs?.totalDirectCostAllCohorts ?? (s.costs?.directCostPerCohort ?? 0) * (c.cohorts || 0), 0);
      }
      if (h.includes("direct cost") && h.includes("per cohort")) {
        return cellCurrency(s.costs?.directCostPerCohort ?? 0, 0);
      }

      if (h.includes("economic cost") && (h.includes("total") || h.includes("all cohorts") || h.includes("national"))) {
        return cellCurrency(s.costs?.totalEconomicCostAllCohorts ?? s.natTotalCost ?? 0, 0);
      }
      if (h.includes("economic cost") && h.includes("per cohort")) {
        return cellCurrency(s.costs?.totalEconomicCostPerCohort ?? 0, 0);
      }

      if (h.includes("preference model")) return `<td>${safeText(s.preferenceModel)}</td>`;
      if (h.includes("endorse")) return cellNumeric(s.endorseRate, 1);

      if ((h.includes("wtp") || h.includes("perceived")) && h.includes("per trainee")) return cellCurrency(s.wtpPerTraineePerMonth, 0);
      if ((h.includes("wtp") || h.includes("perceived")) && h.includes("total") && h.includes("cohort")) return cellCurrency(s.wtpPerCohort, 0);
      if ((h.includes("wtp") || h.includes("perceived")) && (h.includes("all cohorts") || h.includes("national"))) return cellCurrency(s.wtpAllCohorts, 0);

      if (h.includes("bcr") && (h.includes("cohort") || h.includes("per cohort"))) {
        return `<td class="numeric-cell">${s.bcrPerCohort !== null ? formatNumber(s.bcrPerCohort, 2) : "-"}</td>`;
      }
      if (h.includes("bcr") && (h.includes("national") || h.includes("all cohorts"))) {
        return `<td class="numeric-cell">${s.natBcr !== null ? formatNumber(s.natBcr, 2) : "-"}</td>`;
      }

      if (h.includes("total economic cost")) return cellCurrency(s.natTotalCost, 0);
      if (h.includes("epi benefit") && (h.includes("all cohorts") || h.includes("national"))) return cellCurrency(s.epiBenefitAllCohorts, 0);
      if (h.includes("epi benefit")) return cellCurrency(s.epiBenefitPerCohort, 0);
      if (h.includes("net benefit") && (h.includes("all cohorts") || h.includes("national"))) return cellCurrency(s.netBenefitAllCohorts, 0);
      if (h.includes("net benefit")) return cellCurrency(s.netBenefitPerCohort, 0);

      if (h.includes("feasibility") || (h.includes("capacity") && h.includes("status"))) return `<td>${safeText(cap.status)}</td>`;
      if (h.includes("mentor shortfall")) return cellNumeric(cap.mentorShortfall, 0);

      if (h.includes("actions") || h.includes("view")) {
        return `<td><button class="mini" data-snapshot="${s.id}">View</button></td>`;
      }

      // Fallback: keep blank cell to preserve alignment
      return `<td></td>`;
    };

    const tr = document.createElement("tr");

    if (headers.length > 0) {
      tr.innerHTML = headers.map((h) => cellsByHeader(h)).join("");
    } else {
      // Backwards compatible fallback (should not happen)
      tr.innerHTML = `
        <td><input type="checkbox" data-scenario-id="${s.id}" aria-label="Shortlist scenario"></td>
        <td>${safeText(c.name)}</td>
        ${tagsHtml}
        <td>${tierLabel}</td>
        <td>${careerLabel}</td>
        <td>${mentorshipLabel}</td>
        <td>${deliveryLabel}</td>
        <td>${responseLabel}</td>
        <td class="numeric-cell">${formatNumber(c.cohorts, 0)}</td>
        <td class="numeric-cell">${formatNumber(c.traineesPerCohort, 0)}</td>
        <td class="numeric-cell">${formatCurrencyDisplay(c.costPerTraineePerMonth, 0)}</td>
        <td>${safeText(s.preferenceModel)}</td>
        <td class="numeric-cell">${formatNumber(s.endorseRate, 1)}%</td>
        <td class="numeric-cell">${formatCurrencyDisplay(s.wtpPerTraineePerMonth, 0)}</td>
        <td class="numeric-cell">${formatCurrencyDisplay(s.wtpAllCohorts, 0)}</td>
        <td class="numeric-cell">${formatCurrencyDisplay(s.natTotalCost, 0)}</td>
        <td class="numeric-cell">${formatCurrencyDisplay(s.epiBenefitAllCohorts, 0)}</td>
        <td class="numeric-cell">${formatCurrencyDisplay(s.netBenefitAllCohorts, 0)}</td>
        <td class="numeric-cell">${s.natBcr !== null ? formatNumber(s.natBcr, 2) : "-"}</td>
        <td><button class="mini" data-snapshot="${s.id}">View</button></td>
      `;
    }

    tbody.appendChild(tr);
  });

  // Attach snapshot click handler (delegated)
  tbody.querySelectorAll('button[data-snapshot]').forEach((btn) => {
    btn.addEventListener("click", (e) => {
      const id = e.currentTarget.getAttribute("data-snapshot");
      const sc = appState.savedScenarios.find((x) => x.id === id);
      if (sc) openSnapshotModal(sc);
    });
  });

  // Persist shortlist selections on scenario objects (used by Top options panel and exports)
  tbody.querySelectorAll('input[type="checkbox"][data-scenario-id]').forEach((cb) => {
    cb.addEventListener("change", (e) => {
      const id = e.currentTarget.getAttribute("data-scenario-id");
      const sc = appState.savedScenarios.find((x) => x.id === id);
      if (sc) sc.shortlisted = !!e.currentTarget.checked;
      refreshTopScenariosPanel("top5");
  refreshTopScenariosPanel("top5-copilot");
    });
  });

  // Pin/unpin scenarios (used by Copilot prompt generation and quick comparison)
  tbody.querySelectorAll('input[type="checkbox"][data-pin-id]').forEach((cb) => {
    cb.addEventListener("change", (e) => {
      const id = e.currentTarget.getAttribute("data-pin-id");
      const sc = appState.savedScenarios.find((x) => x.id === id);
      if (sc) sc.pinned = !!e.currentTarget.checked;
      refreshTopScenariosPanel("top5");
      refreshTopScenariosPanel("top5-copilot");
      if (typeof window.renderCopilotPromptBundle === "function") {
        window.renderCopilotPromptBundle();
      }
    });
  });

  // Update the Top options panel whenever the table is rebuilt
  refreshTopScenariosPanel("top5");
  refreshTopScenariosPanel("top5-copilot");
}

/* ===========================
   Top options (saved scenarios)
   =========================== */

function computeScenarioRankScore(s, metric) {
  if (!s) return -Infinity;
  const nb = Number(s.netBenefitAllCohorts || 0);
  const bcr = s.natBcr === null || s.natBcr === undefined ? null : Number(s.natBcr);
  const end = Number(s.endorseRate || 0);

  if (metric === "netBenefit") return nb;
  if (metric === "bcr") return bcr === null ? -Infinity : bcr;
  if (metric === "endorsement") return end;

  // Balanced: normalise within the saved set (robust to scale)
  // We compute a simple z-like scaled score on three dimensions: net benefit, BCR, endorsement.
  // This is a heuristic ranking aid only; it does not change any calculations.
  const all = appState.savedScenarios || [];
  const nbs = all.map((x) => Number(x.netBenefitAllCohorts || 0));
  const bcrs = all
    .map((x) => (x.natBcr === null || x.natBcr === undefined ? null : Number(x.natBcr)))
    .filter((x) => x !== null && isFinite(x));
  const ends = all.map((x) => Number(x.endorseRate || 0));

  function minMax(arr) {
    const finite = arr.filter((v) => isFinite(v));
    if (!finite.length) return { min: 0, max: 1 };
    return { min: Math.min(...finite), max: Math.max(...finite) };
  }

  const r1 = minMax(nbs);
  const r2 = minMax(bcrs.length ? bcrs : [0]);
  const r3 = minMax(ends);

  function scaled(v, r) {
    if (!isFinite(v)) return 0;
    const denom = r.max - r.min;
    if (denom <= 0) return 0.5;
    return clamp((v - r.min) / denom, 0, 1);
  }

  const s_nb = scaled(nb, r1);
  const s_bcr = scaled(bcr === null ? r2.min : bcr, r2);
  const s_end = scaled(end, r3);

  // Weights (kept simple; can be adjusted later without affecting model outputs)
  return 0.45 * s_nb + 0.35 * s_bcr + 0.20 * s_end;
}


function isScenarioShortlisted(s) {
  return !!(s && s.shortlisted);
}

function getScenarioDisplayName(s) {
  if (!s) return "Scenario";
  const c = s.config || {};
  const name = (c.name || s.name || "").trim();
  if (name) return name;
  // Defensive fallback: rebuild a concise name from config if available
  try {
    return buildDefaultScenarioName({
      tier: c.tier,
      career: c.career,
      mentorship: c.mentorship,
      delivery: c.delivery,
      cohorts: c.cohorts,
      traineesPerCohort: c.traineesPerCohort,
      costPerTraineePerMonth: c.costPerTraineePerMonth
    });
  } catch (e) {
    return "Scenario";
  }
}

function computeTopScenarioItems(metric, feasibleOnly, selectedOnly, limit) {
  const candidates = (appState.savedScenarios || []).map((s, idx) => ({ s, idx }))
    .filter((x) => (selectedOnly ? isScenarioShortlisted(x.s) : true))
    .map((x) => {
      const cap = x.s.capacity || computeCapacity(x.s.config || {});
      const feasible = cap.status === "Within current capacity";
      return { ...x, cap, feasible };
    })
    .filter((x) => (feasibleOnly ? x.feasible : true));

  if (!candidates.length) return [];

  if (metric === "balanced") {
    const nbs = candidates.map((x) => Number(x.s.netBenefitAllCohorts || 0));
    const bcrs = candidates
      .map((x) => (x.s.natBcr === null || x.s.natBcr === undefined ? NaN : Number(x.s.natBcr)))
      .filter((v) => isFinite(v));
    const ends = candidates.map((x) => Number(x.s.endorseRate || 0));

    function minMax(arr) {
      const finite = arr.filter((v) => isFinite(v));
      if (!finite.length) return { min: 0, max: 1 };
      return { min: Math.min(...finite), max: Math.max(...finite) };
    }
    const r1 = minMax(nbs);
    const r2 = minMax(bcrs.length ? bcrs : [0]);
    const r3 = minMax(ends);

    function scaled(v, r) {
      if (!isFinite(v)) return 0;
      const denom = r.max - r.min;
      if (denom <= 0) return 0.5;
      return clamp((v - r.min) / denom, 0, 1);
    }

    candidates.forEach((x) => {
      const nb = Number(x.s.netBenefitAllCohorts || 0);
      const bcr = x.s.natBcr === null || x.s.natBcr === undefined ? NaN : Number(x.s.natBcr);
      const end = Number(x.s.endorseRate || 0);
      const s_nb = scaled(nb, r1);
      const s_bcr = isFinite(bcr) ? scaled(bcr, r2) : 0;
      const s_end = scaled(end, r3);
      x.score = 0.45 * s_nb + 0.35 * s_bcr + 0.20 * s_end;
    });
  }

// Non-balanced ranking: sort by the requested metric with robust handling of missing values.
function keyFor(x) {
  if (metric === "bcr") {
    const v = Number(x.s.natBcr);
    return isFinite(v) ? v : -Infinity;
  }
  if (metric === "endorsement") {
    const v = Number(x.s.endorseRate);
    return isFinite(v) ? v : -Infinity;
  }
  const v = Number(x.s.netBenefitAllCohorts);
  return isFinite(v) ? v : -Infinity;
}

return candidates
  .slice()
  .sort((a, b) => {
    const kb = keyFor(b);
    const ka = keyFor(a);
    if (kb === ka) return a.idx - b.idx; // stable tie break
    return kb - ka;
  })
  .slice(0, limit || 5);

}

function setScenarioShortlistByIndex(idx, value) {
  if (idx === null || idx === undefined) return;
  const s = appState.savedScenarios[idx];
  if (!s) return;
  s.shortlisted = !!value;
  refreshSavedScenariosTable();
  refreshSensitivityTables();
}

function renderTopScenariosTable(prefix, items) {
  const body = document.getElementById(`${prefix}-table-body`);
  const emptyNote = document.getElementById(`${prefix}-empty`);
  if (!body) return;

  body.innerHTML = "";

  if (!items || items.length === 0) {
    if (emptyNote) emptyNote.classList.remove("hidden");
    return;
  }
  if (emptyNote) emptyNote.classList.add("hidden");

  const isCopilot = prefix === "top5-copilot";

  items.forEach((item, rankIdx) => {
    const s = item.s;
    const cap = item.cap;
    const idx = item.idx;
    const name = getScenarioDisplayName(s);

    const feasibleBadge = cap.status === "Within current capacity"
      ? `<span class="badge badge-ok">Within capacity</span>`
      : `<span class="badge badge-warn">Requires expansion</span>`;

    if (isCopilot) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td class="numeric-cell">${rankIdx + 1}</td>
        <td>${safeText(name)}</td>
        <td class="numeric-cell">${formatPct(s.endorseRate)}</td>
        <td class="numeric-cell">${formatCurrency(s.totalEconomicCostAllCohorts || 0)}</td>
        <td class="numeric-cell">${formatCurrency(s.netBenefitAllCohorts || 0)}</td>
        <td>${feasibleBadge}</td>
      `;
      body.appendChild(tr);
      return;
    }

    const isShort = isScenarioShortlisted(s);
    const shortlistLabel = isShort ? "Unshortlist" : "Shortlist";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="numeric-cell">${rankIdx + 1}</td>
      <td>${safeText(name)}</td>
      <td class="numeric-cell">${formatPct(s.endorseRate)}</td>
      <td class="numeric-cell">${formatCurrency(s.perceivedValuePerTraineePerMonth || s.wtpPerTraineePerMonth || 0)}</td>
      <td class="numeric-cell">${formatCurrency(s.totalEconomicCostAllCohorts || 0)}</td>
      <td class="numeric-cell">${formatCurrency(s.totalEpiBenefitsAllCohorts || 0)}</td>
      <td class="numeric-cell">${formatCurrency(s.netBenefitAllCohorts || 0)}</td>
      <td class="numeric-cell">${formatNumber(s.natBcr, 2)}</td>
      <td>${feasibleBadge}</td>
      <td class="numeric-cell">${formatNumber(cap.mentorShortfall || 0, 0)}</td>
      <td>
        <div class="row-actions">
          <button type="button" class="btn-link btn-link-small" data-top5-action="view" data-sidx="${idx}">View</button>
          <button type="button" class="btn-link btn-link-small" data-top5-action="shortlist" data-sidx="${idx}">${shortlistLabel}</button>
        </div>
      </td>
    `;
    body.appendChild(tr);
  });

  // Attach delegated listeners once per render
  if (!isCopilot) {
    body.querySelectorAll("[data-top5-action='view']").forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = Number(btn.getAttribute("data-sidx"));
        const s = appState.savedScenarios[idx];
        if (s) openSnapshotModal(s);
      });
    });
    body.querySelectorAll("[data-top5-action='shortlist']").forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = Number(btn.getAttribute("data-sidx"));
        const s = appState.savedScenarios[idx];
        if (!s) return;
        setScenarioShortlistByIndex(idx, !isScenarioShortlisted(s));
      });
    });
  }
}

function refreshTopScenariosPanel(prefix = "top5") {
  const panel = document.getElementById(`${prefix}-panel`);
  const body = document.getElementById(`${prefix}-table-body`);
  if (!panel || !body) return;

  const rankByEl = document.getElementById(`${prefix}-rank-by`);
  const feasibleOnlyEl = document.getElementById(`${prefix}-feasible-only`);
  const selectedOnlyEl = document.getElementById(`${prefix}-selected-only`);

  const metric = rankByEl ? rankByEl.value : "netBenefit";
  const feasibleOnly = feasibleOnlyEl ? feasibleOnlyEl.checked : false;
  const selectedOnly = selectedOnlyEl ? selectedOnlyEl.checked : false;

  const items = computeTopScenarioItems(metric, feasibleOnly, selectedOnly, 5);
  renderTopScenariosTable(prefix, items);
}

// Backwards compatibility: old name used elsewhere
function refreshTopScenariosPanelLegacy() {
  refreshTopScenariosPanel("top5");
}


function exportScenariosToExcel() {
  if (!window.XLSX) {
    showToast("Excel export is not available in this browser.", "error");
    return;
  }

  const wb = XLSX.utils.book_new();

  // Sheet 1: Scenario summary table
  const rows = [];
  rows.push([
    "Name",
    "Tier",
    "Career",
    "Mentorship",
    "Delivery",
    "Response time (days)",
    "Cohorts",
    "Trainees per cohort",
    "Cost per trainee per month (INR)",
    "Opportunity cost included",
    "Mentor support cost base per cohort (INR)",
    "Mentor multiplier",
    "Mentor support cost per cohort (INR)",
    "Total mentor support cost all cohorts (INR)",
    "Programme cost per cohort (INR)",
    "Direct cost per cohort (INR)",
    "Opportunity cost per cohort (INR)",
    "Salary-based opportunity cost per cohort (additional, INR)",
    "Salary-based OC faculty (INR)",
    "Salary-based OC coordinator (INR)",
    "Salary-based OC participant (INR)",
    "Salary monthly faculty (INR)",
    "Salary monthly coordinator (INR)",
    "Salary monthly participant (INR)",
    "Economic cost per cohort (INR)",
    "Total direct cost all cohorts (INR)",
    "Total economic cost all cohorts (INR)",
    "Preference model",
    "Endorsement (%)",
    "Perceived programme value per trainee per month (INR)",
    "Perceived programme value per cohort (INR)",
    "Total perceived programme value all cohorts (INR)",
    "Total epidemiological benefit per cohort (INR)",
    "Total epidemiological benefit all cohorts (INR)",
    "Net epidemiological benefit per cohort (INR)",
    "Net epidemiological benefit all cohorts (INR)",
    "BCR per cohort",
    "BCR national",
    "Outbreak responses per year (national)",
    "Cross-sector benefit multiplier",
    "Mentors required per cohort",
    "Total mentors required nationally",
    "Available mentors nationally",
    "Mentor shortfall",
    "Feasibility status",
    "Available training sites / hubs",
    "Max cohorts per site per year",
    "Implied annual site capacity (cohorts)",
    "Site capacity gap (cohorts)",
    "Notes",
    "Enablers (export)",
    "Risks (export)"
  ]);

  appState.savedScenarios.forEach((s) => {
    const c = s.config;
    const cap = s.capacity || computeCapacity(c);
    const a = buildAssumptionsForScenario(s);

    rows.push([
      c.name || "Scenario",
      c.tier,
      c.career,
      c.mentorship,
      c.delivery,
      c.response,
      c.cohorts,
      c.traineesPerCohort,
      c.costPerTraineePerMonth,
      c.opportunityCostIncluded ? "Yes" : "No",
      a.mentorSupportCostPerCohortBase,
      a.mentorMultiplierApplied,
      s.costs?.mentorCostPerCohort ?? 0,
      s.costs?.totalMentorCostAllCohorts ?? (s.costs?.mentorCostPerCohort ?? 0) * (c.cohorts || 0),
      s.costs?.programmeCostPerCohort ?? 0,
      s.costs?.directCostPerCohort ?? 0,
      s.costs?.opportunityCostPerCohort ?? 0,
      s.costs?.salaryBasedOpportunityCostPerCohort ?? 0,
      (s.costs?.salaryBasedOpportunityCostByRole?.faculty ?? 0),
      (s.costs?.salaryBasedOpportunityCostByRole?.coordinator ?? 0),
      (s.costs?.salaryBasedOpportunityCostByRole?.participant ?? 0),
      (c.salaryMonthlyFaculty ?? SALARY_OC_DEFAULTS.facultyMonthly),
      (c.salaryMonthlyCoordinator ?? SALARY_OC_DEFAULTS.coordinatorMonthly),
      (c.salaryMonthlyParticipant ?? SALARY_OC_DEFAULTS.participantMonthly),
      s.costs?.totalEconomicCostPerCohort ?? 0,
      s.costs?.totalDirectCostAllCohorts ?? (s.costs?.directCostPerCohort ?? 0) * (c.cohorts || 0),
      s.costs?.totalEconomicCostAllCohorts ?? s.natTotalCost ?? 0,
      s.preferenceModel,
      s.endorseRate,
      s.wtpPerTraineePerMonth,
      s.wtpPerCohort,
      s.wtpAllCohorts,
      s.epiBenefitPerCohort,
      s.epiBenefitAllCohorts,
      s.netBenefitPerCohort,
      s.netBenefitAllCohorts,
      s.bcrPerCohort !== null ? s.bcrPerCohort : "",
      s.natBcr !== null ? s.natBcr : "",
      s.outbreaksPerYearNational,
      a.crossSectorBenefitMultiplier,
      cap.mentorsPerCohort,
      cap.totalMentorsRequired,
      cap.availableMentors,
      cap.mentorShortfall,
      cap.status,
      a.availableTrainingSites,
      a.maxCohortsPerSitePerYear,
      cap.siteCapacity !== null ? cap.siteCapacity : "",
      cap.siteGap !== null ? cap.siteGap : "",
      c.notes || "",
      c.exportEnablers || "",
      c.exportRisks || ""
    ]);
  });

  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, sheet, "STEPS scenarios");

  // Sheet 2: Assumptions used (scenario-specific)
  const aRows = [];
  aRows.push([
    "Scenario",
    "Planning horizon (years)",
    "Discount rate (%)",
    "Completion rate (%)",
    "Outbreak responses per graduate per year",
    "Value per outbreak (INR)",
    "Non-outbreak value per graduate per year (INR)",
    "Opportunity cost included",
    "Mentor cost base per cohort (INR)",
    "Mentorship multiplier applied",
    "Cross-sector benefit multiplier",
    "Available mentors nationally",
    "Available training sites / hubs",
    "Max cohorts per site per year",
    "Salary monthly faculty (INR)",
    "Salary monthly coordinator (INR)",
    "Salary monthly participant (INR)",
    "Salary OC contact days (frontline)",
    "Salary OC contact days (intermediate)",
    "Salary OC contact days (advanced)"
  ]);

  appState.savedScenarios.forEach((s) => {
    const a = buildAssumptionsForScenario(s);
    aRows.push([
      s.config.name || "Scenario",
      a.planningHorizonYears,
      a.discountRate * 100,
      a.completionRate * 100,
      a.outbreaksPerGraduatePerYear,
      a.valuePerOutbreak,
      a.valuePerGraduate,
      a.opportunityCostIncluded ? "Yes" : "No",
      a.mentorSupportCostPerCohortBase,
      a.mentorMultiplierApplied,
      a.crossSectorBenefitMultiplier,
      a.availableMentorsNational,
      a.availableTrainingSites,
      a.maxCohortsPerSitePerYear,
      a.salaryMonthlyFaculty,
      a.salaryMonthlyCoordinator,
      a.salaryMonthlyParticipant,
      a.salaryOcContactDaysFrontline,
      a.salaryOcContactDaysIntermediate,
      a.salaryOcContactDaysAdvanced
    ]);
  });

  const aSheet = XLSX.utils.aoa_to_sheet(aRows);
  XLSX.utils.book_append_sheet(wb, aSheet, "Assumptions");

  XLSX.writeFile(wb, "steps_saved_scenarios.xlsx");
  showToast("Excel file downloaded.", "success");
}



/* ===========================
   PDF rendering helpers (no external plugins)
   =========================== */

function initPdfDoc(orientation = "portrait") {
  if (!window.jspdf || !window.jspdf.jsPDF) return null;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation, unit: "pt", format: "a4" });

  // Baseline styling (World Bank / WHO-like neutrality)
  doc.setFont("helvetica", "normal");
  doc.setTextColor(33, 37, 41);
  doc.setDrawColor(210, 215, 220);
  doc.setLineWidth(0.6);
  return doc;
}

function addPdfPageNumbers(doc, margin = 36) {
  const n = doc.getNumberOfPages ? doc.getNumberOfPages() : 1;
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  doc.setFontSize(9);
  doc.setTextColor(120, 128, 136);
  for (let i = 1; i <= n; i++) {
    doc.setPage(i);
    doc.text(`Page ${i} of ${n}`, pageW - margin, pageH - margin / 2, { align: "right" });
  }
  doc.setTextColor(33, 37, 41);
}

function pdfSplit(doc, text, width) {
  return doc.splitTextToSize(String(text ?? ""), width);
}

function pdfComputeColumnWidths(doc, head, body, tableW) {
  const cols = head.length;
  const maxLens = new Array(cols).fill(0);
  const sampleRows = body.slice(0, 18);

  for (let c = 0; c < cols; c++) {
    maxLens[c] = Math.max(maxLens[c], String(head[c] || "").length);
    sampleRows.forEach((row) => {
      maxLens[c] = Math.max(maxLens[c], String(row[c] || "").length);
    });
    // Cap to avoid a single long cell dominating width
    maxLens[c] = Math.min(maxLens[c], 42);
  }

  // Convert char length to width approximation, then scale to fit
  const raw = maxLens.map((l) => 22 + l * 3.4);
  const sum = raw.reduce((a, b) => a + b, 0) || 1;
  const scaled = raw.map((w) => (w / sum) * tableW);

  // Enforce minimums
  const minW = 58;
  for (let c = 0; c < cols; c++) scaled[c] = Math.max(minW, scaled[c]);

  // Re-scale if we exceeded available width due to minimums
  const sum2 = scaled.reduce((a, b) => a + b, 0) || 1;
  const factor = tableW / sum2;
  return scaled.map((w) => w * factor);
}

function pdfDrawTable(doc, opts) {
  const {
    x,
    y,
    w,
    head,
    body,
    fontSize = 8.5,
    headFontSize = 8.5,
    lineHeight = 10.5,
    cellPadding = 4,
    zebra = true
  } = opts;

  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const marginBottom = 36;

  const colW = pdfComputeColumnWidths(doc, head, body, w);

  function ensureSpace(hNeed) {
    if (y + hNeed <= pageH - marginBottom) return;
    doc.addPage();
    y = 36;
  }

  // Header row
  ensureSpace(lineHeight + 8);

  doc.setFontSize(headFontSize);
  doc.setFont("helvetica", "bold");
  doc.setFillColor(245, 247, 250);
  doc.rect(x, y, w, lineHeight + cellPadding, "F");
  doc.setDrawColor(210, 215, 220);
  doc.rect(x, y, w, lineHeight + cellPadding);

  let cx = x;
  for (let c = 0; c < head.length; c++) {
    const txt = pdfSplit(doc, head[c], colW[c] - 2 * cellPadding);
    doc.text(txt.slice(0, 2), cx + cellPadding, y + lineHeight * 0.85);
    cx += colW[c];
    if (c < head.length - 1) doc.line(cx, y, cx, y + lineHeight + cellPadding);
  }
  y += lineHeight + cellPadding;

  // Body rows
  doc.setFont("helvetica", "normal");
  doc.setFontSize(fontSize);

  for (let r = 0; r < body.length; r++) {
    const row = body[r];
    // determine wrapped lines per cell
    const wrapped = row.map((val, c) => pdfSplit(doc, val, colW[c] - 2 * cellPadding));
    const maxLines = Math.max(1, ...wrapped.map((arr) => arr.length));
    const rowH = maxLines * lineHeight + cellPadding;

    ensureSpace(rowH + 1);

    if (zebra && r % 2 === 1) {
      doc.setFillColor(252, 253, 254);
      doc.rect(x, y, w, rowH, "F");
    }

    doc.setDrawColor(210, 215, 220);
    doc.rect(x, y, w, rowH);

    cx = x;
    for (let c = 0; c < head.length; c++) {
      const lines = wrapped[c].slice(0, 5); // avoid runaway rows
      doc.text(lines, cx + cellPadding, y + lineHeight);
      cx += colW[c];
      if (c < head.length - 1) doc.line(cx, y, cx, y + rowH);
    }
    y += rowH;
  }

  return y;
}


function parseBulletsText(text) {
  const raw = String(text || "").trim();
  if (!raw) return [];
  let parts = raw.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
  if (parts.length === 1) {
    const semi = parts[0].split(";").map((x) => x.trim()).filter(Boolean);
    if (semi.length > 1) parts = semi;
  }
  return parts;
}

function pdfDrawBullets(doc, bullets, x, y, width) {
  const fontSize = doc.getFontSize ? doc.getFontSize() : 10;
  const lh = Math.round(fontSize * 1.35);
  const indent = 10;

  bullets.forEach((b) => {
    const lines = pdfSplit(doc, String(b), width - indent, fontSize);
    if (!lines.length) return;
    doc.text("•", x, y);
    doc.text(lines[0], x + indent, y);
    y += lh;
    for (let i = 1; i < lines.length; i++) {
      doc.text(lines[i], x + indent, y);
      y += lh;
    }
    y += 2;
  });

  return y;
}

function metricLabel(metric) {
  switch (metric) {
    case "bcr":
      return "BCR";
    case "endorsement":
      return "Endorsement";
    case "balanced":
      return "Balanced score (transparent)";
    case "netBenefit":
    default:
      return "Net benefit";
  }
}

// Simple text box helper for PDFs (no external plugins)
function pdfDrawTextBox(doc, opts) {
  const x = opts.x ?? 40;
  const y = opts.y ?? 40;
  const width = opts.width ?? 500;
  const text = String(opts.text ?? "");
  const fontSize = opts.fontSize ?? 10;
  const padding = opts.padding ?? 10;
  const lineHeight = opts.lineHeight ?? Math.round(fontSize * 1.35);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(fontSize);

  const innerW = Math.max(20, width - 2 * padding);
  const lines = pdfSplit(doc, text, innerW, fontSize);

  const h = padding * 2 + lines.length * lineHeight;

  doc.setDrawColor(200);
  doc.setFillColor(248, 249, 251);
  doc.roundedRect(x, y, width, h, 6, 6, "FD");

  let ty = y + padding + fontSize;
  lines.forEach((ln) => {
    doc.text(ln, x + padding, ty);
    ty += lineHeight;
  });

  return y + h;
}


/* ===== Word and Excel export helpers (docx/xlsx) ===== */
function getDocxLib() {
  return (window.docx || (typeof docx !== "undefined" ? docx : null));
}

function formatDateForFilename(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function triggerDownload(blob, filename) {
  try {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    try { a.click(); } catch (eClick) {}
    try { a.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); } catch (eEvt) {}
    setTimeout(() => {
      URL.revokeObjectURL(url);
      a.remove();
    }, 500);
  } catch (e) {
    showToast("Download failed in this browser. Please try again or use a different browser.", "error");
  }
}

function downloadTextFile(filename, text) {
  const blob = new Blob([String(text || "")], { type: "text/plain;charset=utf-8" });
  triggerDownload(blob, filename);
}


async function downloadDocx(doc, filename) {
  const lib = getDocxLib();
  if (!lib || !lib.Packer) {
    showToast("Word export is unavailable because the Word generation library failed to load.", "error");
    return false;
  }
  try {
    const blob = await lib.Packer.toBlob(doc);
    triggerDownload(blob, filename);
    return true;
  } catch (e) {
    showToast("Failed to generate the Word document. Please try again.", "error");
    return false;
  }
}

function docxParagraph(text, opts) {
  const lib = getDocxLib();
  if (!lib) return null;
  const { Paragraph, TextRun, HeadingLevel, AlignmentType } = lib;
  const options = opts || {};
  if (options.heading) {
    return new Paragraph({ text: text || "", heading: options.heading, spacing: { after: options.after ?? 120 } });
  }
  if (options.bold || options.italic) {
    return new Paragraph({
      children: [new TextRun({ text: text || "", bold: !!options.bold, italic: !!options.italic })],
      spacing: { after: options.after ?? 120 }
    });
  }
  return new Paragraph({ text: text || "", spacing: { after: options.after ?? 120 }, alignment: options.alignment || undefined });
}

function docxKeyValueTable(rows) {
  const lib = getDocxLib();
  if (!lib) return null;
  const { Table, TableRow, TableCell, Paragraph, WidthType } = lib;
  const tableRows = rows.map((r) => new TableRow({
    children: [
      new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: String(r[0] ?? ""), spacing: { after: 0 } })] }),
      new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: String(r[1] ?? ""), spacing: { after: 0 } })] })
    ]
  }));
  return new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } });
}

function docxScenarioComparisonTable(scenarios, baseline) {
  const lib = getDocxLib();
  if (!lib) return null;
  const { Table, TableRow, TableCell, Paragraph, WidthType } = lib;

  const header = new TableRow({
    children: [
      new TableCell({ children: [new Paragraph({ text: "Scenario", spacing: { after: 0 } })] }),
      new TableCell({ children: [new Paragraph({ text: "Endorsement", spacing: { after: 0 } })] }),
      new TableCell({ children: [new Paragraph({ text: "Economic cost (INR)", spacing: { after: 0 } })] }),
      new TableCell({ children: [new Paragraph({ text: "Epidemiological benefits (INR)", spacing: { after: 0 } })] }),
      new TableCell({ children: [new Paragraph({ text: "Net benefit (INR)", spacing: { after: 0 } })] }),
      new TableCell({ children: [new Paragraph({ text: "BCR", spacing: { after: 0 } })] }),
      new TableCell({ children: [new Paragraph({ text: "Feasibility", spacing: { after: 0 } })] })
    ]
  });

  const rows = [];
  if (baseline) {
    rows.push(new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ text: "Baseline (status quo)", spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: "-", spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: formatCurrencyDisplay(baseline.natTotalCost || 0, 0), spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: formatCurrencyDisplay(baseline.epiBenefitAllCohorts || 0, 0), spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: formatCurrencyDisplay(baseline.netBenefitAllCohorts || 0, 0), spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: baseline.natBcr != null ? formatNumber(baseline.natBcr, 2) : "-", spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: "Baseline parameters", spacing: { after: 0 } })] })
      ]
    }));
  }

  (scenarios || []).forEach((s) => {
    const cap = s.capacity || computeCapacity(s.config || {});
    const feas = cap.status === "Within current capacity" ? "Within current capacity" : "Requires capacity expansion";
    rows.push(new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ text: getScenarioDisplayName(s), spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: formatPct(s.endorseRate), spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: formatCurrencyDisplay(s.totalEconomicCostAllCohorts || 0, 0), spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: formatCurrencyDisplay(s.totalEpiBenefitsAllCohorts || 0, 0), spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: formatCurrencyDisplay(s.netBenefitAllCohorts || 0, 0), spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: s.natBcr != null ? formatNumber(s.natBcr, 2) : "-", spacing: { after: 0 } })] }),
        new TableCell({ children: [new Paragraph({ text: feas, spacing: { after: 0 } })] })
      ]
    }));
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [header, ...rows]
  });
}
function exportScenariosToPdf() {
  // Word export that works reliably in-browser (including file://) by downloading an HTML-based Word document.
  const mode = getExportMode();
  const notes = getExportNotesFromUI();
  const scenarios = (mode === "brief") ? getShortlistedOrTopScenarios(3) : (Array.isArray(appState.savedScenarios) ? appState.savedScenarios.slice() : []);
  if (!scenarios || scenarios.length === 0) {
    showToast("Save at least one scenario before exporting.", "warning");
    return;
  }

  const baseline = appState.baselineScenario || getBaselineForCurrentHorizon();
  const now = new Date();
  const dateLabel = now.toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });

  const refreshed = (Array.isArray(scenarios) ? scenarios : []).map((s) => refreshScenarioForExport(s)).filter(Boolean);
  const seen = new Set();
  const unique = refreshed.filter((s) => {
    const key = s._sid || s.id || getScenarioDisplayName(s);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  const sorted = unique.slice().sort((a, b) => (Number(b.netBenefitAllCohorts || 0) - Number(a.netBenefitAllCohorts || 0)));
  const best = sorted[0];
  const cap = best ? (best.capacity || computeCapacity(best.config || {})) : null;
  const feasText = cap ? (cap.status || "") : "";

  const title = "STEPS policy brief and scenario comparison";
  const intro =
    "This document summarises the baseline (business as usual) and compares saved scale up scenarios created in STEPS. " +
    "Economic cost is the sum of direct financial cost plus two opportunity cost components: the existing opportunity cost and the salary-based opportunity cost (additional). " +
    "The salary-based component reflects time diverted from regular duties for faculty, coordinators and participants, and is not a cash payment.";

  const body = [
    `<h1>${safeText(title)}</h1>`,
    `<p class="small muted">Generated on ${safeText(dateLabel)}.</p>`,
    `<p>${safeText(intro)}</p>`,
    `<h2>Baseline and scenarios</h2>`,
    `<p>Net benefit is epidemiological benefit minus economic cost. BCR is epidemiological benefit divided by economic cost.</p>`,
    scenarioComparisonHtmlTable(sorted, baseline),
    `<h2>Graduate outputs and economic results</h2>`,
    `<p>This table restates the core outputs that drive the benefit–cost results. It shows the expected number of graduates by tier and the corresponding economic cost, epidemiological benefit, net benefit and benefit–cost ratio. All values reflect the full planning horizon and the number of cohorts configured for each scenario.</p>`,
    scenarioOutputsHtmlTable(sorted, baseline),
    (best ? `<h2>Interpretation in plain language</h2><p>The highest net benefit scenario in this export is <strong>${safeText(getScenarioDisplayName(best))}</strong>. Under the current assumptions, feasibility is assessed as ${safeText(feasText.toLowerCase() || "not available")}. If decision makers change benefit inputs in the Settings or Sensitivity tabs, regenerate this report to keep totals internally consistent. When reading the numbers, remember that economic cost includes both direct spending and time diverted from routine duties (opportunity cost). The epidemiological benefit reflects the value of improved outbreak detection and response capacity as parameterised in STEPS.</p>` : ""),
    (formatExportNotes(notes) ? `<h2>Notes added in STEPS</h2>${formatExportNotes(notes)}` : "")
  ].join("\n");

  const html = wordDocWrapHtml(title, body);
  downloadWordDoc(`STEPS_policy_brief_${formatDateForFilename(now)}`, html);
}

function exportTop5OnlyPdf() {
  // Word export that works reliably in-browser (including file://) by downloading an HTML-based Word document.

  const rankByEl = document.getElementById("top5-rank-by");
  const feasibleEl = document.getElementById("top5-feasible-only");
  const selectedEl = document.getElementById("top5-selected-only");

  const metric = rankByEl ? rankByEl.value : "netBenefit";
  const feasibleOnly = feasibleEl ? feasibleEl.checked : false;
  const selectedOnly = selectedEl ? selectedEl.checked : false;

  const items = computeTopScenarioItems(metric, feasibleOnly, selectedOnly, 5);
  if (!items || items.length === 0) {
    showToast("Save scenarios first to export the Top 5 list.", "warning");
    return;
  }

  const scenarios = items.map((it) => appState.savedScenarios[it.idx]).filter(Boolean).map((s) => refreshScenarioForExport(s));
  const baseline = appState.baselineScenario || getBaselineForCurrentHorizon();
  const now = new Date();
  const dateLabel = now.toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });

  const metricLabel = metric === "bcr" ? "benefit cost ratio" : metric === "endorsement" ? "endorsement" : metric === "balanced" ? "balanced score" : "net benefit";
  const intro =
    `This Top 5 list is generated from scenarios saved in STEPS. Ranking is by ${metricLabel}. ` +
    `Feasible only filter is ${feasibleOnly ? "on" : "off"}. Shortlisted only filter is ${selectedOnly ? "on" : "off"}. ` +
    `This ranking does not change any model calculations. It only sorts saved scenarios.`;

  const body = [
    `<h1>STEPS Top 5 scenarios</h1>`,
    `<p class="small muted">Generated on ${safeText(dateLabel)}.</p>`,
    `<p>${safeText(intro)}</p>`,
    scenarioComparisonHtmlTable(scenarios, baseline),
    `<h2>Graduate outputs and economic results</h2>`,
    `<p>This table helps interpret the Top 5 list by showing graduates by tier and the associated economic cost and epidemiological benefit.</p>`,
    scenarioOutputsHtmlTable(scenarios, baseline)
  ].join("\n");

  const html = wordDocWrapHtml("STEPS Top 5 scenarios", body);
  downloadWordDoc(`STEPS_top5_${formatDateForFilename(now)}`, html);
}

function wordDocWrapHtml(title, bodyHtml) {
  const safeTitle = String(title || "STEPS export").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const css = `
    @page { margin: 2.2cm; }
    body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #111; line-height: 1.35; }
    h1 { font-size: 18pt; margin: 0 0 10pt 0; }
    h2 { font-size: 14pt; margin: 14pt 0 6pt 0; }
    h3 { font-size: 12pt; margin: 12pt 0 6pt 0; }
    p { margin: 0 0 8pt 0; }
    .small { font-size: 10pt; color: #444; }
    table { border-collapse: collapse; width: 100%; margin: 8pt 0 12pt 0; }
    th, td { border: 1px solid #D0D7DE; padding: 6pt; vertical-align: top; }
    th { background: #F3F4F6; font-weight: 700; }
    .num { text-align: right; white-space: nowrap; }
    .muted { color: #555; }
    .note { border: 1px solid #D0D7DE; background: #F8FAFC; padding: 10pt; margin: 10pt 0 12pt 0; }
  `.trim();

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>${safeTitle}</title>
<style>${css}</style>
</head>
<body>
${bodyHtml || ""}
</body>
</html>
  `.trim();
}

function scenarioComparisonHtmlTable(scenarios, baseline) {
  const rows = (Array.isArray(scenarios) ? scenarios : []).map((s) => {
    const name = safeText(getScenarioDisplayName(s));
    const cap = s.capacity || computeCapacity(s.config || {});
    const cost = Number(s.natTotalCost || s.totalEconomicCostAllCohorts || 0);
    const ben = Number(s.epiBenefitAllCohorts || s.totalEpiBenefitsAllCohorts || 0);
    const net = (s.netBenefitAllCohorts != null) ? Number(s.netBenefitAllCohorts) : (ben - cost);
    const bcr = (cost > 0) ? (ben / cost) : 0;
    const end = (s.endorseRate != null) ? formatPct(s.endorseRate) : "";
    return `
      <tr>
        <td>${name}</td>
        <td class="num">${end}</td>
        <td class="num">${safeText(formatCurrencyDisplay(cost, 0))}</td>
        <td class="num">${safeText(formatCurrencyDisplay(ben, 0))}</td>
        <td class="num">${safeText(formatCurrencyDisplay(net, 0))}</td>
        <td class="num">${formatNumber(bcr, 2)}</td>
        <td>${safeText(cap.status || "")}</td>
      </tr>
    `.trim();
  }).join("\n");

  const bCost = Number(baseline?.natTotalCost || 0);
  const bBen = Number(baseline?.epiBenefitAllCohorts || 0);
  const bNet = Number(baseline?.netBenefitAllCohorts || 0);
  const bBcr = (bCost > 0) ? (bBen / bCost) : 0;

  const baselineRow = `
    <tr>
      <td><strong>Baseline (business as usual)</strong></td>
      <td class="num">${formatPct(baseline?.endorseRate || 0)}</td>
      <td class="num">${safeText(formatCurrencyDisplay(bCost, 0))}</td>
      <td class="num">${safeText(formatCurrencyDisplay(bBen, 0))}</td>
      <td class="num">${safeText(formatCurrencyDisplay(bNet, 0))}</td>
      <td class="num">${formatNumber(bBcr, 2)}</td>
      <td>${safeText((baseline?.capacity && baseline.capacity.status) ? baseline.capacity.status : "")}</td>
    </tr>
  `.trim();

  return `
    <table>
      <thead>
        <tr>
          <th>Scenario</th>
          <th class="num">Endorsement</th>
          <th class="num">Economic cost (INR)</th>
          <th class="num">Epidemiological benefit (INR)</th>
          <th class="num">Net benefit (INR)</th>
          <th class="num">BCR</th>
          <th>Feasibility</th>
        </tr>
      </thead>
      <tbody>
        ${baselineRow}
        ${rows}
      </tbody>
    </table>
    <p class="small muted">Notes: Economic cost is direct financial spending plus the existing opportunity cost plus the salary-based opportunity cost (additional). The salary-based component reflects time diverted from normal duties and is not a cash payment.</p>
  `.trim();
}


function scenarioOutputsHtmlTable(scenarios, baseline) {
  const rows = (Array.isArray(scenarios) ? scenarios : []).map((s) => {
    const name = safeText(getScenarioDisplayName(s));
    const grads = deriveGraduatesByTierForExport(s);
    const total = grads.total;
    const f = grads.frontline;
    const i = grads.intermediate;
    const a = grads.advanced;
    const cost = safeNumber(s.natTotalCost ?? s.totalEconomicCostAllCohorts ?? 0, 0);
    const ben = safeNumber(s.epiBenefitAllCohorts ?? s.totalEpiBenefitsAllCohorts ?? 0, 0);
    const net = (s.netBenefitAllCohorts != null) ? safeNumber(s.netBenefitAllCohorts, 0) : (ben - cost);
    const bcr = (cost > 0) ? (ben / cost) : 0;
    return `
      <tr>
        <td>${name}</td>
        <td class="num">${formatNumber(total, 0)}</td>
        <td class="num">${formatNumber(f, 0)}</td>
        <td class="num">${formatNumber(i, 0)}</td>
        <td class="num">${formatNumber(a, 0)}</td>
        <td class="num">${safeText(formatCurrencyDisplay(cost, 0))}</td>
        <td class="num">${safeText(formatCurrencyDisplay(ben, 0))}</td>
        <td class="num">${safeText(formatCurrencyDisplay(net, 0))}</td>
        <td class="num">${formatNumber(bcr, 2)}</td>
      </tr>
    `.trim();
  }).join("\n");

  const bTotal = safeNumber(baseline?.natTotalGrads ?? 0, 0);
  const bF = safeNumber(baseline?.natFrontlineGrads ?? 0, 0);
  const bI = safeNumber(baseline?.natIntermediateGrads ?? 0, 0);
  const bA = safeNumber(baseline?.natAdvancedGrads ?? 0, 0);
  const bCost = safeNumber(baseline?.natTotalCost ?? 0, 0);
  const bBen = safeNumber(baseline?.epiBenefitAllCohorts ?? 0, 0);
  const bNet = safeNumber(baseline?.netBenefitAllCohorts ?? (bBen - bCost), 0);
  const bBcr = (bCost > 0) ? (bBen / bCost) : 0;

  const baselineRow = `
    <tr>
      <td><strong>Baseline (business as usual)</strong></td>
      <td class="num">${formatNumber(bTotal, 0)}</td>
      <td class="num">${formatNumber(bF, 0)}</td>
      <td class="num">${formatNumber(bI, 0)}</td>
      <td class="num">${formatNumber(bA, 0)}</td>
      <td class="num">${safeText(formatCurrencyDisplay(bCost, 0))}</td>
      <td class="num">${safeText(formatCurrencyDisplay(bBen, 0))}</td>
      <td class="num">${safeText(formatCurrencyDisplay(bNet, 0))}</td>
      <td class="num">${formatNumber(bBcr, 2)}</td>
    </tr>
  `.trim();

  return `
    <table>
      <thead>
        <tr>
          <th>Scenario</th>
          <th class="num">Total graduates</th>
          <th class="num">Frontline</th>
          <th class="num">Intermediate</th>
          <th class="num">Advanced</th>
          <th class="num">Economic cost (INR)</th>
          <th class="num">Epidemiological benefit (INR)</th>
          <th class="num">Net benefit (INR)</th>
          <th class="num">BCR</th>
        </tr>
      </thead>
      <tbody>
        ${baselineRow}
        ${rows}
      </tbody>
    </table>
  `.trim();
}


function downloadWordDoc(filenameBase, html) {
  const name = String(filenameBase || "steps_export").replace(/[^a-z0-9_\-]+/gi, "_");
  const filename = name.toLowerCase().endsWith(".doc") ? name : `${name}.doc`;
  const content = String(html || "");

  // Primary method: Blob + object URL
  try {
    const blob = new Blob([content], { type: "application/msword;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    try { a.click(); } catch (eClick) {}
    try { a.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); } catch (eEvt) {}
    setTimeout(() => {
      try { URL.revokeObjectURL(url); } catch (e) {}
      try { a.remove(); } catch (e) {}
    }, 0);
    return;
  } catch (e) {
    // fall through
  }

  // Fallback: data URI download
  try {
    const uri = "data:application/msword;charset=utf-8," + encodeURIComponent(content);
    const a = document.createElement("a");
    a.href = uri;
    a.download = filename;
    a.target = "_blank";
    a.rel = "noopener";
    a.style.display = "none";
    document.body.appendChild(a);
    try { a.click(); } catch (eClick) {}
    try { a.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); } catch (eEvt) {}
    a.remove();
    return;
  } catch (e2) {
    // fall through
  }

  // Final fallback: open a new window with the document content.
  try {
    const w = window.open("", "_blank");
    if (w && w.document) {
      w.document.open();
      w.document.write(content);
      w.document.close();
    }
  } catch (e3) {
    // ignore
  }
}

function sanitizePara(text) {
  const raw = String(text || "").trim();
  if (!raw) return "";
  // Remove any accidental bullet prefixes and em dashes
  const cleaned = raw
    .replace(/^\s*[\-\u2022]\s+/gm, "")
    .replace(/Not provided/g, "-")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned;
}

function splitIntoParagraphs(text) {
  const raw = String(text || "");
  const parts = raw.split(/\n+/).map((t) => sanitizePara(t)).filter(Boolean);
  return parts.length ? parts : [];
}

function buildScenarioSummaryTableHtml(scenarios, baselineScenario) {
  const rows = (scenarios || []).map((s) => {
    const name = escapeHtml(s.config && s.config.name ? s.config.name : "Scenario");
    const tier = escapeHtml(getTierLabel(s.config && s.config.tier));
    const delivery = escapeHtml(getDeliveryLabel(s.config && s.config.delivery));
    const cohorts = safeNumber(s.config && s.config.cohorts, 0);
    const trainees = safeNumber(s.config && s.config.traineesPerCohort, 0);
    const endorse = safeNumber(s.endorseRate, 1);
    const econCost = safeNumber(s.costs && s.costs.totalEconomicCostAllCohortsINR, 0);
    const epiBenefit = safeNumber(s.epiBenefitAllCohorts || 0, 0);
    const netBenefit = safeNumber(s.netBenefitAllCohorts || 0, 0);
    const bcr = safeNumber(s.bcrAllCohorts || 0, 2);
    const feasible = escapeHtml((s.capacity && s.capacity.feasible) ? "Feasible" : "Not feasible");

    const existingOpp = safeNumber(s.costs && s.costs.existingOpportunityCostAllCohortsINR, 0);
    const salaryOpp = safeNumber(s.costs && s.costs.salaryOpportunityCostAllCohortsINR, 0);

    const baseline = baselineScenario || null;
    const baseCost = baseline ? (baseline.costs && baseline.costs.totalEconomicCostAllCohortsINR) : null;
    const baseBenefit = baseline ? (baseline.epiBenefitAllCohorts || 0) : null;
    const baseNet = baseline ? (baseline.netBenefitAllCohorts || 0) : null;

    const dCost = (baseCost !== null && baseCost !== undefined) ? (econCost - baseCost) : null;
    const dBenefit = (baseBenefit !== null && baseBenefit !== undefined) ? (epiBenefit - baseBenefit) : null;
    const dNet = (baseNet !== null && baseNet !== undefined) ? (netBenefit - baseNet) : null;

    const deltaCost = (dCost === null) ? "Not available" : formatSignedCurrency(dCost);
    const deltaBenefit = (dBenefit === null) ? "Not available" : formatSignedCurrency(dBenefit);
    const deltaNet = (dNet === null) ? "Not available" : formatSignedCurrency(dNet);

    return `
      <tr>
        <td>${name}</td>
        <td>${tier}</td>
        <td>${delivery}</td>
        <td class="num">${formatNumber(cohorts, 0)}</td>
        <td class="num">${formatNumber(trainees, 0)}</td>
        <td class="num">${formatNumber(endorse, 1)}%</td>
        <td class="num">${formatCurrencyDisplay(econCost, 0)}</td>
        <td class="num">${formatCurrencyDisplay(existingOpp, 0)}</td>
        <td class="num">${formatCurrencyDisplay(salaryOpp, 0)}</td>
        <td class="num">${formatCurrencyDisplay(epiBenefit, 0)}</td>
        <td class="num">${formatCurrencyDisplay(netBenefit, 0)}</td>
        <td class="num">${formatNumber(bcr, 2)}</td>
        <td>${feasible}</td>
        <td class="num">${deltaCost}</td>
        <td class="num">${deltaBenefit}</td>
        <td class="num">${deltaNet}</td>
      </tr>
    `.trim();
  }).join("\n");

  return `
    <table>
      <thead>
        <tr>
          <th>Scenario</th>
          <th>Tier</th>
          <th>Delivery</th>
          <th class="num">Cohorts</th>
          <th class="num">Trainees per cohort</th>
          <th class="num">Endorsement</th>
          <th class="num">Economic cost (total)</th>
          <th class="num">Existing opportunity cost</th>
          <th class="num">Salary-based opportunity cost</th>
          <th class="num">Epidemiological benefit (total)</th>
          <th class="num">Net benefit</th>
          <th class="num">Benefit-cost ratio</th>
          <th>Feasibility</th>
          <th class="num">Change in cost vs baseline</th>
          <th class="num">Change in benefit vs baseline</th>
          <th class="num">Change in net benefit vs baseline</th>
        </tr>
      </thead>
      <tbody>
        ${rows || ""}
      </tbody>
    </table>
  `.trim();
}

function buildScenarioInterpretationsHtml(scenarios, baselineScenario) {
  const baseline = baselineScenario || null;

  return (scenarios || []).map((s) => {
    const name = escapeHtml(s.config && s.config.name ? s.config.name : "Scenario");
    const tier = escapeHtml(getTierLabel(s.config && s.config.tier));
    const delivery = escapeHtml(getDeliveryLabel(s.config && s.config.delivery));
    const endorse = safeNumber(s.endorseRate, 1);
    const econCost = safeNumber(s.costs && s.costs.totalEconomicCostAllCohortsINR, 0);
    const epiBenefit = safeNumber(s.epiBenefitAllCohorts || 0, 0);
    const netBenefit = safeNumber(s.netBenefitAllCohorts || 0, 0);
    const bcr = safeNumber(s.bcrAllCohorts || 0, 2);
    const feasible = (s.capacity && s.capacity.feasible) ? "feasible within the current mentor and hub inputs" : "not feasible without expanding mentors and or hubs";

    let deltaText = "";
    if (baseline) {
      const dCost = econCost - (baseline.costs && baseline.costs.totalEconomicCostAllCohortsINR || 0);
      const dBenefit = epiBenefit - (baseline.epiBenefitAllCohorts || 0);
      const dNet = netBenefit - (baseline.netBenefitAllCohorts || 0);
      deltaText = `Compared with the baseline, this option changes total economic cost by ${formatSignedCurrency(dCost)}, changes total epidemiological benefit by ${formatSignedCurrency(dBenefit)}, and changes net benefit by ${formatSignedCurrency(dNet)}.`;
    } else {
      deltaText = "Baseline comparison is not available because a baseline has not been saved in the Planner tab.";
    }

    const existingOpp = safeNumber(s.costs && s.costs.existingOpportunityCostAllCohortsINR, 0);
    const salaryOpp = safeNumber(s.costs && s.costs.salaryOpportunityCostAllCohortsINR, 0);

    const paragraph1 = `This scenario provides a ${tier} programme delivered as ${delivery}. Estimated endorsement is ${formatNumber(endorse, 1)} percent. Total economic cost is ${formatCurrencyDisplay(econCost, 0)} and includes existing opportunity cost of ${formatCurrencyDisplay(existingOpp, 0)} plus salary-based opportunity cost of ${formatCurrencyDisplay(salaryOpp, 0)}.`;
    const paragraph2 = `Total epidemiological benefit is ${formatCurrencyDisplay(epiBenefit, 0)}. Net benefit is ${formatCurrencyDisplay(netBenefit, 0)} with a benefit-cost ratio of ${formatNumber(bcr, 2)}. This option is ${feasible}.`;
    const paragraph3 = deltaText;

    return `
      <h3>${name}</h3>
      <p>${escapeHtml(paragraph1)}</p>
      <p>${escapeHtml(paragraph2)}</p>
      <p>${escapeHtml(paragraph3)}</p>
    `.trim();
  }).join("\n");
}

function formatSignedCurrency(value) {
  const v = safeNumber(value, 0);
  const sign = v >= 0 ? "+" : "-";
  return `${sign}${formatCurrencyDisplay(Math.abs(v), 0)}`;
}



/* ===========================
   WTP based benefits and sensitivity
   =========================== */

function getSensitivityControls() {
  const benefitModeSelect = getElByIdCandidates(["benefit-definition-select", "benefitDefinitionSelect"]);
  const epiToggle = getElByIdCandidates(["sensitivity-epi-toggle", "sensitivityEpiToggle"]);
  const endorsementOverrideInput = getElByIdCandidates(["endorsement-override", "endorsementOverride"]);

  return {
    benefitMode: benefitModeSelect ? benefitModeSelect.value : "wtp_only",
    epiIncluded: epiToggle && epiToggle.classList.contains("on"),
    endorsementOverride: endorsementOverrideInput ? Number(endorsementOverrideInput.value) || null : null
  };
}

function computeSensitivityRow(scenario) {
  const c = scenario.config;
  const costAll = scenario.costs.totalEconomicCostPerCohort * c.cohorts;
  const epiAll = scenario.epiBenefitPerCohort * c.cohorts;
  const netAll = epiAll - costAll;
  const epiBcr = costAll > 0 ? epiAll / costAll : null;

  const wtpAll = scenario.wtpAllCohorts;
  const wtpOutbreak = scenario.wtpOutbreakComponent;
  const combinedBenefit = wtpAll + epiAll;

  const npvDceOnly = wtpAll - costAll;
  const npvCombined = combinedBenefit - costAll;

  return { costAll, epiAll, netAll, epiBcr, wtpAll, wtpOutbreak, combinedBenefit, npvDceOnly, npvCombined };
}

function refreshSensitivityTables() {
  const dceBody = document.getElementById("dce-benefits-table-body");
  const sensBody = document.getElementById("sensitivity-table-body");
  if (!dceBody || !sensBody) return;

  dceBody.innerHTML = "";
  sensBody.innerHTML = "";

  if (!appState.currentScenario) return;

  const controls = getSensitivityControls();

  const scenarios = [
    { label: "Current configuration", scenario: appState.currentScenario },
    ...appState.savedScenarios.map((s, idx) => ({
      label: s.config.name || `Saved scenario ${idx + 1}`,
      scenario: s
    }))
  ];

  scenarios.forEach(({ label, scenario }) => {
    const c = scenario.config;
    const s = computeSensitivityRow(scenario);

    let endorsementUsed = controls.endorsementOverride !== null ? controls.endorsementOverride : scenario.endorseRate;
    endorsementUsed = clamp(endorsementUsed, 0, 100);

         let effectiveWtp = s.wtpAll;
      if (controls.benefitMode === "endorsement_adjusted") {
        effectiveWtp = s.wtpAll * (endorsementUsed / 100);
      }

      // Combined benefit for sensitivity analysis is defined using perceived programme value only.
      // Epidemiological outbreak benefits are reported in a separate column and not added into the BCR/NPV here.
      let combinedBenefit = effectiveWtp;

      const bcrDceOnly = s.costAll > 0 ? s.wtpAll / s.costAll : null;
      const bcrCombined = s.costAll > 0 ? combinedBenefit / s.costAll : null;

      const npvDceOnly = s.npvDceOnly;
      const npvCombined = combinedBenefit - s.costAll;

    const trHeadline = document.createElement("tr");
        trHeadline.innerHTML = `
          <td>${label}</td>
          <td class="numeric-cell">${formatCurrencyINR(s.costAll, 0)}</td>
          <td class="numeric-cell">${formatNumber(s.costAll / 1e6, 2)}</td>
          <td class="numeric-cell">${formatNumber(s.netAll / 1e6, 2)}</td>
          <td class="numeric-cell">${formatCurrencyINR(s.wtpAll, 0)}</td>
          <td class="numeric-cell">${formatCurrencyINR(s.wtpOutbreak, 0)}</td>
          <td class="numeric-cell">${formatNumber(endorsementUsed, 1)}</td>
          <td class="numeric-cell">${formatCurrencyINR(effectiveWtp, 0)}</td>
          <td class="numeric-cell">${bcrDceOnly !== null ? formatNumber(bcrDceOnly, 2) : "-"}</td>
          <td class="numeric-cell">${formatCurrencyINR(npvDceOnly, 0)}</td>
          <td class="numeric-cell">${bcrCombined !== null ? formatNumber(bcrCombined, 2) : "-"}</td>
          <td class="numeric-cell">${formatCurrencyINR(npvCombined, 0)}</td>
        `;
        dceBody.appendChild(trHeadline);

        const trDetail = document.createElement("tr");
        trDetail.innerHTML = `
          <td>${label}</td>
          <td>${scenario.preferenceModel}</td>
          <td class="numeric-cell">${formatNumber(scenario.endorseRate, 1)}%</td>
          <td class="numeric-cell">${formatCurrencyINR(scenario.costs.totalEconomicCostPerCohort, 0)}</td>
          <td class="numeric-cell">${formatCurrencyINR(scenario.wtpPerCohort, 0)}</td>
          <td class="numeric-cell">${formatCurrencyINR(scenario.wtpOutbreakComponent, 0)}</td>
          <td class="numeric-cell">${bcrDceOnly !== null ? formatNumber(bcrDceOnly, 2) : "-"}</td>
          <td class="numeric-cell">${formatCurrencyINR(npvDceOnly / c.cohorts, 0)}</td>
          <td class="numeric-cell">${bcrCombined !== null ? formatNumber(bcrCombined, 2) : "-"}</td>
          <td class="numeric-cell">${formatCurrencyINR(npvCombined / c.cohorts, 0)}</td>
          <td class="numeric-cell">${formatCurrencyINR((s.wtpAll * (endorsementUsed / 100)) / c.cohorts, 0)}</td>
        `;
        sensBody.appendChild(trDetail);
  });
}

function exportSensitivityToExcel() {
  if (!window.XLSX) {
    showToast("Excel export is not available in this browser.", "error");
    return;
  }
  const table = document.getElementById("dce-benefits-table");
  if (!table) return;

  const wb = XLSX.utils.book_new();
  const sheet = XLSX.utils.table_to_sheet(table);
  XLSX.utils.book_append_sheet(wb, sheet, "Sensitivity");
  XLSX.writeFile(wb, "steps_sensitivity_summary.xlsx");
  showToast("Sensitivity table Excel file downloaded.", "success");
}

/* ===========================
   Sensitivity contract controls and PDF export
   =========================== */

function exportSensitivityContainerToPdf() {
  // PDF downloads have been replaced by Word document downloads (.doc).
  const table = document.getElementById("sensitivity-table");
  if (!table) {
    showToast("Sensitivity table is not available.", "error");
    return;
  }

  const now = new Date();
  const dateStr = now.toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });

  const intro = `
    <h1>STEPS sensitivity results</h1>
    <p class="small">Generated on ${escapeHtml(dateStr)} from the STEPS decision aid tool.</p>
    <p>This table shows how scenario results change when the selected benefit value is adjusted. Values are calculated using the same scenarios and assumptions as in the tool view.</p>
  `;

  // Use the displayed table as the source of truth
  const tableHtml = `<h2>Sensitivity table</h2>${table.outerHTML}`;

  const html = wordDocWrapHtml("STEPS sensitivity results", `${intro}${tableHtml}`);
  downloadWordDoc("steps_sensitivity_table", html);
  showToast("Word document downloaded.", "success");
}




function initSensitivityContractControls() {
  const select = getElByIdCandidates(["sensitivityValueSelect", "sensitivity-value-select", "sensitivity-value"]);
  const applyBtn = getElByIdCandidates(["applySensitivityValueBtn", "apply-sensitivity-value", "applySensitivityBtn"]);
  const pdfBtn = getElByIdCandidates(["downloadSensitivityPDF", "download-sensitivity-pdf", "downloadSensitivityPdf"]);

  if (select) {
    ensureSelectHasOutbreakPresets(select);
    select.addEventListener("change", () => {
      const valueInINR = parseSensitivityValueToINR(select.value);
      if (valueInINR) {
        applyOutbreakPreset(valueInINR, { silentToast: true, silentLog: true });
      }
    });
  }

  if (applyBtn) {
    applyBtn.addEventListener("click", () => {
      if (!select) {
        showToast("Sensitivity value selector is not available.", "error");
        return;
      }
      const valueInINR = parseSensitivityValueToINR(select.value);
      if (!valueInINR) {
        showToast("Select a valid sensitivity value before applying.", "warning");
        return;
      }
      applyOutbreakPreset(valueInINR, { silentToast: true, silentLog: false });
      showToast("Sensitivity value applied.", "success");
    });
  }

  if (pdfBtn) {
    pdfBtn.addEventListener("click", () => {
      exportSensitivityContainerToPdf();
    });
  }
}

/* ===========================
   Advanced settings
   =========================== */

function logSettingsMessage(message) {
  const targets = [];
  const sessionLog = document.getElementById("settings-log");
  const advLog = document.getElementById("adv-settings-log");
  const contractLog = document.getElementById("settingsLog");

  if (contractLog) targets.push(contractLog);
  if (sessionLog && sessionLog !== contractLog) targets.push(sessionLog);
  if (advLog && advLog !== sessionLog && advLog !== contractLog) targets.push(advLog);

  if (!targets.length) return;

  const time = new Date().toLocaleString();
  targets.forEach((box) => {
    const p = document.createElement("p");
    p.textContent = `[${time}] ${message}`;
    box.appendChild(p);
  });
}

function initAdvancedSettings() {
  const valueGradInput = document.getElementById("adv-value-per-graduate");
  const valueOutbreakInput = document.getElementById("adv-value-per-outbreak");
  const completionInput = document.getElementById("adv-completion-rate");
  const outbreaksPerGradInput = document.getElementById("adv-outbreaks-per-graduate");
  const horizonInput = document.getElementById("adv-planning-horizon");
  const discInput = document.getElementById("adv-epi-discount-rate");
  const usdRateInput = document.getElementById("adv-usd-rate");
  const applyBtn = document.getElementById("adv-apply-settings");
  const resetBtn = document.getElementById("adv-reset-settings");

  function writeLog(message) {
    logSettingsMessage(message);
  }

  if (applyBtn) {
    applyBtn.addEventListener("click", () => {
      if (valueGradInput && valueOutbreakInput && completionInput && outbreaksPerGradInput && horizonInput && discInput && usdRateInput) {
        const vGrad = Number(valueGradInput.value);
        const vOutParsed = parseSensitivityValueToINR(valueOutbreakInput.value);
        const vOut = vOutParsed !== null ? vOutParsed : Number(valueOutbreakInput.value);
        const compRateRaw = Number(completionInput.value);
        const compRate = isFinite(compRateRaw) ? clamp(compRateRaw / 100, 0, 1) : appState.epiSettings.tiers.frontline.completionRate;
        const outPerGrad = Number(outbreaksPerGradInput.value);
        const horizon = Number(horizonInput.value);
        const discRateRaw = Number(discInput.value);
        const discRate = isFinite(discRateRaw) ? clamp(discRateRaw / 100, 0, 1) : appState.epiSettings.general.epiDiscountRate;
        const usdRate = Number(usdRateInput.value);

        ["frontline", "intermediate", "advanced"].forEach((tier) => {
          appState.epiSettings.tiers[tier].valuePerGraduate = isFinite(vGrad) ? vGrad : 0;
          if (isFinite(vOut) && vOut > 0) appState.epiSettings.tiers[tier].valuePerOutbreak = vOut;
          appState.epiSettings.tiers[tier].completionRate = compRate;
          if (isFinite(outPerGrad) && outPerGrad >= 0) appState.epiSettings.tiers[tier].outbreaksPerGraduatePerYear = outPerGrad;
        });

        if (isFinite(horizon) && horizon > 0) appState.epiSettings.general.planningHorizonYears = horizon;
        appState.epiSettings.general.epiDiscountRate = discRate;

        if (isFinite(usdRate) && usdRate > 0) {
          appState.epiSettings.general.inrToUsdRate = usdRate;
          appState.usdRate = usdRate;
        }

        writeLog(
          "Advanced settings updated for graduate value, value per outbreak, completion rate, outbreaks per graduate, planning horizon, discount rate and INR per USD. Current outbreak cost saving calculations use the outbreak value and planning horizon."
        );

        syncOutbreakValueDropdownsFromState();

        if (appState.currentScenario) {
          const newScenario = computeScenario(appState.currentScenario.config);
          appState.currentScenario = newScenario;
          refreshAllOutputs(newScenario);
        }

        showToast("Advanced settings applied for this session.", "success");
      }
    });
  }

  if (resetBtn) {
    resetBtn.addEventListener("click", () => {
      appState.epiSettings = JSON.parse(JSON.stringify(DEFAULT_EPI_SETTINGS));
      appState.usdRate = DEFAULT_EPI_SETTINGS.general.inrToUsdRate;

      if (valueGradInput) valueGradInput.value = "0";
      if (valueOutbreakInput) valueOutbreakInput.value = "4000000";
      if (completionInput) completionInput.value = "90";
      if (outbreaksPerGradInput) outbreaksPerGradInput.value = "0.5";
      if (horizonInput) horizonInput.value = String(DEFAULT_EPI_SETTINGS.general.planningHorizonYears);
      if (discInput) discInput.value = String(DEFAULT_EPI_SETTINGS.general.epiDiscountRate * 100);
      if (usdRateInput) usdRateInput.value = String(DEFAULT_EPI_SETTINGS.general.inrToUsdRate);

      writeLog("Advanced settings reset to default values.");

      syncOutbreakValueDropdownsFromState();

      if (appState.currentScenario) {
        const newScenario = computeScenario(appState.currentScenario.config);
        appState.currentScenario = newScenario;
        refreshAllOutputs(newScenario);
      }

      showToast("Advanced settings reset to defaults.", "success");
    });
  }
}

function applyOutbreakPreset(valueInINR, options = {}) {
  const silentToast = !!options.silentToast;
  const silentLog = !!options.silentLog;

  if (isNaN(valueInINR) || valueInINR <= 0) return;

  ["frontline", "intermediate", "advanced"].forEach((tier) => {
    appState.epiSettings.tiers[tier].valuePerOutbreak = valueInINR;
  });

  const valueOutbreakInput = document.getElementById("adv-value-per-outbreak");
  if (valueOutbreakInput) valueOutbreakInput.value = String(valueInINR);

  syncOutbreakValueDropdownsFromState();

  if (appState.currentScenario) {
    const newScenario = computeScenario(appState.currentScenario.config);
    appState.currentScenario = newScenario;
    refreshAllOutputs(newScenario);
  }

  if (!silentLog) {
    logSettingsMessage(`Value per outbreak updated to ₹${formatNumber(valueInINR, 0)} per outbreak for all tiers from sensitivity controls.`);
  }

  if (!silentToast) {
    showToast(`Value per outbreak set to ₹${formatNumber(valueInINR, 0)} for all tiers.`, "success");
  }
}

/* ===========================
   Copilot integration
   =========================== */

function buildScenarioJsonForCopilot(scenario) {
  const activeScenario = scenario || appState.currentScenario || computeScenario(getConfigFromForm());
  const baseline = getBaselineForCurrentHorizon();
  const deltasActive = computeIncrementalMetrics(activeScenario, baseline) || {};

  const assumptions = buildAssumptionsBoxData(activeScenario, baseline, deltasActive);

  // Deterministic pinned scenarios ordering
  const pinned = (appState.savedScenarios || [])
    .filter((s) => !!s && !!s.pinned)
    .slice()
    .sort((a, b) => {
      const ak = String(a.id || a._id || a.config?.name || "");
      const bk = String(b.id || b._id || b.config?.name || "");
      return ak.localeCompare(bk);
    });

  const pinnedPayload = pinned.map((s) => {
    const d = computeIncrementalMetrics(s, baseline) || {};
    return {
      id: s.id,
      name: s.config?.name || "Scenario",
      pinned: true,
      config: s.config || {},
      outputs: {
        graduatesAllCohorts: s.graduatesAllCohorts,
        intermediateAdvancedGraduates: getIntermediateAdvancedGraduates(s),
        natTotalCost: s.natTotalCost,
        epiBenefitAllCohorts: s.epiBenefitAllCohorts,
        netBenefitAllCohorts: s.netBenefitAllCohorts,
        natBcr: s.natBcr,
        outbreaksPerYearNational: s.outbreaksPerYearNational,
        wtpAllCohorts: s.wtpAllCohorts
      },
      incremental: d
    };
  });

  return {
    tool: "STEPS",
    version: "planner-baseline",
    timestamp: new Date().toISOString(),
    statusQuoIndia: typeof STATUS_QUO_POLICY_NOTE_TEXT === "string" ? STATUS_QUO_POLICY_NOTE_TEXT : "",
    baseline: {
      inputs: assumptions.baselineInputs,
      outputs: assumptions.outputs.baseline
    },
    activeScenario: {
      id: activeScenario?.id || null,
      name: activeScenario?.config?.name || "Scenario",
      config: activeScenario?.config || {},
      outputs: assumptions.outputs.scenario,
      incremental: deltasActive
    },
    pinnedScenarios: pinnedPayload,
    assumptions: {
      general: assumptions.general
    }
  };
}



function initCopilot() {
  const legacyBtn = document.getElementById("copilot-open-and-copy-btn");

  const copilotBtn = document.getElementById("copilot-copy-btn");
  const chatgptBtn = document.getElementById("chatgpt-copy-btn");
  const copilotRefreshBtn = document.getElementById("copilot-refresh-btn");
  const copilotPolicyBtn = document.getElementById("copilot-copy-policybrief-btn");
  const chatgptPolicyBtn = document.getElementById("chatgpt-copy-policybrief-btn");
  const downloadBtn = document.getElementById("briefing-download-btn");

  const output = document.getElementById("copilot-prompt-output");
  const statusPill = document.getElementById("copilot-status-pill");
  const statusText = document.getElementById("copilot-status-text");

  function setStatus(ok, msg) {
    if (statusPill) statusPill.className = ok ? "pill ok" : "pill warn";
    if (statusText) statusText.textContent = msg;
  }

  function ensureScenarioDataAvailable() {
    // Prompts should work even if no scenarios are saved. If needed, compute a current scenario from the live form inputs.
    try {
      if (appState && appState.currentScenario) return true;
      if (typeof getConfigFromUI === "function" && typeof computeScenario === "function") {
        const cfg = getConfigFromUI();
        const s = computeScenario(cfg);
        if (s && typeof s === "object") {
          if (!s.id) s.id = `current_${Date.now()}`;
          s.config = s.config || cfg;
          appState.currentScenario = s;
          return true;
        }
      }
    } catch (e) {
      // Fall through; we can still generate a template prompt with limited details.
    }
    if (appState && Array.isArray(appState.savedScenarios) && appState.savedScenarios.length) return true;
    showToast("No saved scenarios detected. Prompts will be generated using the current form inputs where possible.", "warning");
    return true;
  }
  function buildPrompt(target) {
    const baseline = getBaselineForCurrentHorizon();
    const saved = Array.isArray(appState.savedScenarios) ? appState.savedScenarios : [];

    // Scenario set: use the National Simulation shortlist if present; otherwise use pinned/shortlisted; otherwise the current scenario.
    const shortlistIdx = Array.isArray(appState.baselineCompareShortlist) ? appState.baselineCompareShortlist : [];
    const shortlist = shortlistIdx
      .map((x) => saved[Number(x)])
      .filter((s) => !!s);

    const pinned = saved.filter((s) => !!s.pinned);
    const shortlisted = saved.filter((s) => !!s.shortlisted);

    const scenarioSet = (shortlist.length ? shortlist :
      (pinned.length || shortlisted.length ? [...pinned, ...shortlisted] : (saved.length ? [saved[0]] : [])))
      .slice(0, 5);

    const currentCfg = (typeof getConfigFromUI === "function") ? getConfigFromUI() : null;
    const currentScenario = (currentCfg && typeof computeScenario === "function") ? computeScenario(currentCfg) : null;

    const scenariosForTables = scenarioSet.length ? scenarioSet : (currentScenario ? [currentScenario] : []);
    const scenarioTable = scenariosForTables.length ? scenarioTableMarkdownRows(scenariosForTables) : "- (no scenario data available)";

    const baselineIAStock = Number(baseline.currentStockIA || 0);
    const targetIAStock = Number(baseline.targetStockIA || 0);
    const baselineGap = Math.max(0, targetIAStock - baselineIAStock);

    const nationalContext = [
      `Baseline target (Intermediate/Advanced): ${formatNumber(targetIAStock,0)}`,
      `Baseline current stock (Intermediate/Advanced): ${formatNumber(baselineIAStock,0)}`,
      `Baseline gap to target: ${formatNumber(baselineGap,0)}`
    ].join("\n");

    const useName = target === "chatgpt" ? "ChatGPT" : "Microsoft Copilot";
    const prompt = [
      `You are drafting a decision note for senior policymakers on FETP India scale-up using outputs from the STEPS tool.`,
      `Write in clear, non-technical language that a Minister can act on immediately. Use short paragraphs, not bullet points.`,
      ``,
      `Use the baseline context and the scenario results table below. Explain what changes, why it matters, and what must be implemented.`,
      ``,
      `Baseline context:`,
      nationalContext,
      ``,
      `Scenario results from STEPS (copy exactly; then interpret):`,
      scenarioTable,
      ``,
      `Your output for ${useName}:`,
      `Produce a 1,200–1,800 word narrative that (i) explains the baseline and the national gap, (ii) compares scenarios, (iii) highlights feasibility constraints, and (iv) recommends a staged implementation plan.`,
      `Include a short table that re-states the key numbers in a Minister-friendly way, and then interpret the numbers in plain language.`,
      ``,
      `Important:`,
      `When discussing costs, distinguish financial cost from economic cost. Economic cost in STEPS includes (1) direct financial cost, (2) the existing opportunity cost module, and (3) the salary-based opportunity cost module (additional, non-cash).`,
      `When discussing benefits, use the epidemiological benefit shown in STEPS and explain in plain terms what drives it.`,
      `If a scenario list is an aggregation of multiple saved scenarios, interpret it as a combined scale-up package (not a single programme).`
    ].join("\n");

    return prompt;
  }

  function buildPolicyBriefPrompt(target) {
    const baseline = getBaselineForCurrentHorizon();
    const saved = Array.isArray(appState.savedScenarios) ? appState.savedScenarios : [];

    // Primary scenario set is the National Simulation shortlist (added list). If empty, use pinned/shortlisted. If still empty, use top ranked.
    const shortlistIdx = Array.isArray(appState.baselineCompareShortlist) ? appState.baselineCompareShortlist : [];
    const shortlist = shortlistIdx.map((x) => saved[Number(x)]).filter((s) => !!s);

    const pinned = saved.filter((s) => !!s.pinned);
    const shortlisted = saved.filter((s) => !!s.shortlisted);

    const rankedMetricEl = document.getElementById("top5-copilot-rank-by");
    const feasibleOnlyEl = document.getElementById("top5-copilot-feasible-only");
    const selectedOnlyEl = document.getElementById("top5-copilot-selected-only");

    const metric = rankedMetricEl ? rankedMetricEl.value : "netBenefit";
    const feasibleOnly = feasibleOnlyEl ? feasibleOnlyEl.checked : false;
    const selectedOnly = selectedOnlyEl ? selectedOnlyEl.checked : true;

    const topItems = computeTopScenarioItems(metric, feasibleOnly, selectedOnly, 5);
    const topScenarios = topItems.map((x) => x.s).filter((s) => !!s);

    const scenarioSet = (shortlist.length ? shortlist :
      (pinned.length || shortlisted.length ? [...pinned, ...shortlisted] :
        (topScenarios.length ? topScenarios : (saved.length ? [saved[0]] : []))))
      .slice(0, 5);

    // Recompute scenarios at prompt-time to ensure any updated inputs are reflected.
    const scenarios = scenarioSet.map((s) => {
      try {
        const cfg = s.config || s.inputs || s._config || null;
        if (cfg && typeof computeScenario === "function") {
          const computed = computeScenario(cfg);
          computed.id = s.id;
          computed.name = s.name || s.title || s.label || getScenarioDisplayName(s);
          computed.shortlisted = s.shortlisted;
          computed.pinned = s.pinned;
          return computed;
        }
      } catch (e) {}
      return s;
    });

    const comparisonTable = scenarios.length ? scenarioTableMarkdownRows(scenarios) : "- (no saved scenarios available)";
    const top5Table = topScenarios.length ? scenarioTableMarkdownRows(topScenarios) : "- (no saved scenarios available)";

    const baselineIAStock = Number(baseline.currentStockIA || 0);
    const targetIAStock = Number(baseline.targetStockIA || 0);
    const baselineGap = Math.max(0, targetIAStock - baselineIAStock);

    const baselineNarrative = [
      `India’s Intermediate and Advanced field epidemiology capacity is benchmarked against an operational target of ${formatNumber(targetIAStock,0)} trained professionals.`,
      `The baseline (business as usual) starting point in STEPS assumes ${formatNumber(baselineIAStock,0)} Intermediate and Advanced trained staff currently available.`,
      `This implies a baseline gap of ${formatNumber(baselineGap,0)} to reach the target.`
    ].join("\n");

    const toolHowTo = [
      `STEPS reports economic cost and epidemiological benefit for each saved scenario. Economic cost includes direct financial cost, the existing opportunity cost module, and the salary-based opportunity cost module (additional).`,
      `Scenario results may be aggregated if the user adds multiple saved scenarios to the National Simulation comparison list; in that case, interpret the Scenario column as the combined package.`
    ].join("\n");

    const useName = target === "chatgpt" ? "ChatGPT" : "Microsoft Copilot";

    const prompt = [
      `You are writing a Minister-ready policy brief (3 to 5 pages) on scaling up Field Epidemiology Training Programmes (FETP) in India under NCDC stewardship, using the STEPS tool outputs provided below.`,
      `Write in clear, non-technical language. Use continuous prose with headings. Avoid bullet points except for a short list of implementation actions at the end, if strictly necessary.`,
      ``,
      `Required output quality: The document must be suitable for immediate policy decision and implementation planning.`,
      ``,
      `Baseline and need for scale-up:`,
      baselineNarrative,
      ``,
      `How STEPS numbers should be interpreted:`,
      toolHowTo,
      ``,
      `Scenario results table (from STEPS). Do not change the numbers:`,
      comparisonTable,
      ``,
      `Top 5 scenarios table (for reference and quick scanning):`,
      top5Table,
      ``,
      `Write the policy brief with the following sections and expectations:`,
      ``,
      `1. Executive summary (about half a page). State the baseline gap, what the recommended package achieves, and the implementation implication.`,
      ``,
      `2. Why scale-up is needed now (about one page). Explain preparedness, surveillance, outbreak response capacity, and why training scale-up is a strategic investment.`,
      ``,
      `3. What the scenarios change compared with baseline (about one to one and a half pages). Use the scenario table. Explain changes in graduates, costs, benefits, net benefits, and benefit–cost ratio in plain language. If more than one scenario is in the list, treat it as a combined package and explain what that means.`,
      ``,
      `4. Cost and benefit interpretation (about one page). Explain financial cost versus economic cost. Clearly explain opportunity costs: (i) the existing module; (ii) the salary-based module (additional) that values time away from routine duties for faculty, coordinators and participants. Explain what drives the epidemiological benefit measure in STEPS.`,
      ``,
      `5. Feasibility and implementation (about one page). Discuss mentors and hub readiness, retention, quality assurance, and monitoring. Identify whether scale-up can be done immediately or requires staged expansion.`,
      ``,
      `6. Recommendation and implementation plan (about half a page). Provide a clear recommended next step, timeline, governance, and minimum monitoring indicators.`,
      ``,
      `Formatting: include at least two clean tables suitable for Word. Use plain headers and aligned columns. Use INR formatting where appropriate. Do not use em dash characters.`,
      ``,
      `Finally, add a short annex (half to one page) explaining what a policymaker should change in STEPS if they want to use a different benefit metric, and what checks must be run to keep results consistent.`,
      ``,
      `This prompt is for ${useName}.`
    ].join("\n");

    return prompt;
  }
  function buildPromptBundle(target) {
    const longPrompt = buildPolicyBriefPrompt(target);
    const analysisPrompt = buildPrompt(target);

    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const d = String(now.getDate()).padStart(2, "0");
    const dateTag = `${y}${m}${d}`;

    const title = `STEPS prompts (${dateTag})`;
    const body = [
      title,
      "",
      "PROMPT A: 3 to 5 page Minister-ready policy brief",
      "",
      longPrompt,
      "",
      "PROMPT B: Shorter analysis and interpretation note",
      "",
      analysisPrompt
    ].join("\n");

    // IMPORTANT: the Copilot prompt panel expects a STRING. Returning an object
    // results in "[object Object]" being copied. Keep the bundle title embedded
    // as the first line of the text so downloads remain meaningful.
    return body;
  }

  
  function buildPromptFallback(target, kind) {
    const targetName = target === "chatgpt" ? "ChatGPT" : "Microsoft Copilot";
    const baseline = appState.baselineScenario || null;
    const current = appState.currentScenario || null;
    const saved = Array.isArray(appState.savedScenarios) ? appState.savedScenarios : [];
    const scenario = current || saved[0] || null;

    const num = (v) => (v === null || v === undefined || !isFinite(Number(v)) ? 0 : Number(v));
    const fmtINR = (v) => (typeof formatINR === "function" ? formatINR(num(v)) : `INR ${num(v).toFixed(0)}`);
    const fmt = (v) => (typeof formatNumber === "function" ? formatNumber(num(v), 2) : num(v).toFixed(2));

    const b = baseline || {};
    const s = scenario || {};
    const bName = b.name || "Baseline (business as usual)";
    const sName = s.name || "Scenario";

    const bCost = num(b.nationalEconomicCost || b.totalEconomicCostAllCohorts || 0);
    const sCost = num(s.nationalEconomicCost || s.totalEconomicCostAllCohorts || 0);

    const bBen = num(b.nationalEpiBenefit || b.epiBenefitAllCohorts || 0);
    const sBen = num(s.nationalEpiBenefit || s.epiBenefitAllCohorts || 0);

    const bNB = num(bBen - bCost);
    const sNB = num(sBen - sCost);

    const bBcr = bCost > 0 ? (bBen / bCost) : 0;
    const sBcr = sCost > 0 ? (sBen / sCost) : 0;

    const rows = [
      `| Indicator | ${bName} | ${sName} | Incremental (Scenario − Baseline) |`,
      `|---|---:|---:|---:|`,
      `| Total economic cost (INR) | ${fmtINR(bCost)} | ${fmtINR(sCost)} | ${fmtINR(sCost - bCost)} |`,
      `| Epidemiological benefit (INR) | ${fmtINR(bBen)} | ${fmtINR(sBen)} | ${fmtINR(sBen - bBen)} |`,
      `| Net benefit (INR) | ${fmtINR(bNB)} | ${fmtINR(sNB)} | ${fmtINR(sNB - bNB)} |`,
      `| Benefit–cost ratio (BCR) | ${fmt(bBcr)} | ${fmt(sBcr)} | ${fmt(sBcr - bBcr)} |`
    ].join("\n");

    const header = kind === "policy"
      ? `Write a 3 to 5 page policy brief for senior decision-makers using outputs from STEPS. Use plain language and continuous prose (avoid bullet points). Include a short executive summary, interpret the numbers, and provide actionable recommendations.`
      : `Use the STEPS outputs below to draft an interpretation-focused briefing note in plain language. Use continuous prose (avoid bullet points).`;

    const guidance = [
      `You are assisting with STEPS (Scalable Training Estimation and Planning System) for FETP India scale-up decisions.`,
      `Use the comparison table to explain what changes in the scenario compared with baseline, what drives costs and benefits, and what the feasibility constraints imply.`,
      `Do not invent inputs. If something is not provided, state that it is not available in the STEPS snapshot.`
    ].join("\n");

    return [
      `TARGET PLATFORM: ${targetName}`,
      ``,
      header,
      ``,
      guidance,
      ``,
      `STEPS comparison (from tool):`,
      rows
    ].join("\n");
  }

  function renderPromptToPanel(target, mode) {
    if (!output) return;
    const t = target || "copilot";
    const m = mode || "bundle";
    let textOut = "";
    try {
      if (m === "policy") textOut = buildPolicyBriefPrompt(t);
      else if (m === "existing") textOut = buildPrompt(t);
      else textOut = buildPromptBundle(t);
    } catch (e) {
      // Fallback prompt that always works, even if a richer builder fails.
      const kind = (m === "policy") ? "policy" : "brief";
      textOut = buildPromptFallback(t, kind);
      if (typeof console !== "undefined" && console && console.warn) console.warn("Copilot prompt fallback used:", e);
    }
    // If the builder returned empty, use fallback as well.
    if (!String(textOut || "").trim()) {
      const kind = (m === "policy") ? "policy" : "brief";
      textOut = buildPromptFallback(t, kind);
    }

    // Defensive: never allow an object to be written into the textarea.
    // (Some earlier builds returned {title, body} for bundles.)
    if (textOut && typeof textOut === "object" && "body" in textOut) {
      textOut = textOut.body;
    }
    output.value = textOut;
    output.dataset.hasGenerated = "true";
    output.dataset.lastTarget = t;
    output.dataset.lastMode = m;
  }

  // Expose a lightweight refresh hook used by horizon/site live updates.
  window.renderCopilotPromptBundle = function () {
    if (!output) return;
    const t = output.dataset.lastTarget || "copilot";
    const m = output.dataset.lastMode || "bundle";
    renderPromptToPanel(t, m);
  };


      
function copyPromptToClipboard(prompt) {
        const text = String(prompt || "");
        // Clipboard API is often unavailable on file://. Use a robust fallback.
        const tryExecCommand = () => {
          try {
            const ta = document.createElement("textarea");
            ta.value = text;
            ta.setAttribute("readonly", "");
            ta.style.position = "absolute";
            ta.style.left = "-9999px";
            document.body.appendChild(ta);
            ta.select();
            ta.setSelectionRange(0, ta.value.length);
            const ok = document.execCommand("copy");
            ta.remove();
            return ok;
          } catch (e) {
            return false;
          }
        };

        if (navigator.clipboard && window.isSecureContext) {
          navigator.clipboard.writeText(text).then(() => {
            setStatus(true, "Prompt copied to clipboard.");
            showToast("Prompt copied to clipboard.", "success");
          }).catch(() => {
            const ok = tryExecCommand();
            if (ok) {
              setStatus(true, "Prompt copied to clipboard.");
              showToast("Prompt copied to clipboard.", "success");
            } else {
              setStatus(false, "Copy failed. Copy manually from the text box.");
              showToast("Copy failed. Please copy from the text box.", "warning");
            }
          });
          return;
        }

        const ok = tryExecCommand();
        if (ok) {
          setStatus(true, "Prompt copied to clipboard.");
          showToast("Prompt copied to clipboard.", "success");
        } else {
          setStatus(false, "Copy failed. Copy manually from the text box.");
          showToast("Copy failed. Please copy from the text box.", "warning");
        }
      }



      function ensureSavedScenariosOrFail() {
        if (!output) return false;

        const hasSaved = !!(appState.savedScenarios && appState.savedScenarios.length > 0);
        const hasCurrent = !!appState.currentScenario;

        if (!hasSaved && !hasCurrent) {
          setStatus(false, "Apply a configuration first.");
          output.value = "";
          showToast("Apply a configuration before generating prompts.", "warning");
          return false;
        }

        if (!hasSaved && hasCurrent) {
          setStatus(true, "Using the currently applied scenario (no saved scenarios).");
          return true;
        }

        return true;
      }

      function handleBundle(target) {
        if (!ensureScenarioDataAvailable()) return;
        renderPromptToPanel(target, "bundle");
        copyPromptToClipboard(output.value);
      }

      function handlePolicy(target) {
        if (!ensureScenarioDataAvailable()) return;
        renderPromptToPanel(target, "policy");
        copyPromptToClipboard(output.value);
      }

      function handleExisting(target) {
        if (!ensureScenarioDataAvailable()) return;
        renderPromptToPanel(target, "existing");
        copyPromptToClipboard(output.value);
      }

      if (copilotBtn) copilotBtn.addEventListener("click", () => handleBundle("copilot"));
      if (chatgptBtn) chatgptBtn.addEventListener("click", () => handleBundle("chatgpt"));
      if (copilotPolicyBtn) copilotPolicyBtn.addEventListener("click", () => handlePolicy("copilot"));
      if (chatgptPolicyBtn) chatgptPolicyBtn.addEventListener("click", () => handlePolicy("chatgpt"));
      if (copilotRefreshBtn) {
        copilotRefreshBtn.addEventListener("click", () => {
          if (!ensureScenarioDataAvailable()) return;
          // Regenerate the prompt using the most up-to-date saved results and National Simulation selection.
          renderPromptToPanel("copilot", "bundle");
          setStatus(true, "Prompt refreshed using the latest results.");
          if (typeof showToast === "function") {
            showToast("Prompt refreshed using the latest saved results and national simulation selection.", "success");
          }
        });
      }

      const copilotOpenBtn = document.getElementById("copilot-open-and-copy-btn");
      const chatgptOpenBtn = document.getElementById("chatgpt-open-and-copy-btn");

      if (copilotOpenBtn) {
        copilotOpenBtn.addEventListener("click", () => {
          if (!ensureScenarioDataAvailable()) return;
        renderPromptToPanel("copilot", "bundle");
          copyPromptToClipboard(output.value);
          window.open("https://copilot.microsoft.com", "_blank", "noopener");
          setStatus(true, "Prompt copied and Copilot opened in a new tab.");
          if (typeof showToast === "function") {
            showToast("Prompt copied and Copilot opened in a new tab.", "success");
          }
        });
      }

      if (chatgptOpenBtn) {
        chatgptOpenBtn.addEventListener("click", () => {
          if (!ensureScenarioDataAvailable()) return;
        renderPromptToPanel("chatgpt", "bundle");
          copyPromptToClipboard(output.value);
          window.open("https://chat.openai.com", "_blank", "noopener");
          setStatus(true, "Prompt copied and ChatGPT opened in a new tab.");
          if (typeof showToast === "function") {
            showToast("Prompt copied and ChatGPT opened in a new tab.", "success");
          }
        });
      }

  if (downloadBtn) {
    downloadBtn.addEventListener("click", () => {
    const promptText = generateAndShowPrompt("copilot","policybrief");
    const dt = new Date();
    const y = dt.getFullYear();
    const mo = String(dt.getMonth()+1).padStart(2,"0");
    const d = String(dt.getDate()).padStart(2,"0");
    const fname = `STEPS_PolicyBrief_Prompt_${y}${mo}${d}.txt`;
    downloadTextFile(fname, promptText);
    showToast("Prompts downloaded.", "success");
  });
    // (legacy download handler removed; download is handled above via downloadTextFile)
  }

  setStatus(true, "Ready.");
  // Generate an initial prompt so the panel is never empty.
  try {
    if (output) {
      ensureScenarioDataAvailable();
      renderPromptToPanel("copilot", "bundle");
      setStatus(true, "Prompt generated.");
    }
  } catch (e) { /* non-fatal */ }
}

  function setPromptText(text) {
    if (!output) return;
    output.value = String(text || "").trim();
    if (!output.value) output.value = "";
  }

  function generateAndShowPrompt(target, variant) {
    ensureScenarioDataAvailable();
    const promptText = buildPrompt(target, variant);
    setPromptText(promptText);
    setStatus(true, "Prompt generated. You can copy or download it.");
    return promptText;
  }

  async function copyText(text) {
    const t = String(text || "");
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(t);
        return true;
      }
    } catch (e) {}
    // Fallback for file://
    try {
      const ta = document.createElement("textarea");
      ta.value = t;
      ta.style.position = "fixed";
      ta.style.left = "-9999px";
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      const ok = document.execCommand("copy");
      document.body.removeChild(ta);
      return ok;
    } catch (e) {
      return false;
    }
  

  if (legacyBtn) {
    legacyBtn.addEventListener("click", async () => {
      const promptText = generateAndShowPrompt("copilot","policybrief");
      const ok = await copyText(promptText);
      window.open("https://copilot.microsoft.com/", "_blank", "noopener");
      showToast(ok ? "Opened Copilot and copied the prompt." : "Opened Copilot. Prompt is ready in the box to copy.", ok ? "success" : "warning");
    });
  }
  if (chatgptOpenBtn) {
    chatgptOpenBtn.addEventListener("click", async () => {
      const promptText = generateAndShowPrompt("chatgpt","policybrief");
      const ok = await copyText(promptText);
      window.open("https://chat.openai.com/", "_blank", "noopener");
      showToast(ok ? "Opened ChatGPT and copied the prompt." : "Opened ChatGPT. Prompt is ready in the box to copy.", ok ? "success" : "warning");
    });
  }
}


/* ===========================
   Copilot prompt buttons fallback wiring
   =========================== */

function copyTextToClipboardRobust(text) {
  const t = String(text || "");
  const tryExecCommand = () => {
    try {
      const ta = document.createElement("textarea");
      ta.value = t;
      ta.setAttribute("readonly", "");
      ta.style.position = "absolute";
      ta.style.left = "-9999px";
      document.body.appendChild(ta);
      ta.select();
      ta.setSelectionRange(0, ta.value.length);
      const ok = document.execCommand("copy");
      ta.remove();
      return ok;
    } catch (e) {
      return false;
    }
  };

  if (navigator.clipboard && window.isSecureContext) {
    return navigator.clipboard.writeText(t).then(() => true).catch(() => tryExecCommand());
  }
  return Promise.resolve(tryExecCommand());
}

function initCopilotPromptButtonFallback() {
  // Capture-phase delegate: ensures the Copilot tab buttons work even if other handlers fail.
  document.addEventListener("click", (ev) => {
    const btn = (ev.target instanceof HTMLElement) ? ev.target.closest("button") : null;
    if (!btn || !btn.id) return;

    const id = btn.id;
    const ids = new Set([
      "copilot-copy-btn",
      "chatgpt-copy-btn",
      "copilot-refresh-btn",
      "copilot-download-btn",
      "copilot-open-and-copy-btn",
      "chatgpt-open-and-copy-btn",
      "copilot-copy-policybrief-btn",
      "chatgpt-copy-policybrief-btn"
    ]);
    if (!ids.has(id)) return;

    ev.preventDefault();
    ev.stopPropagation();
    if (typeof ev.stopImmediatePropagation === "function") ev.stopImmediatePropagation();

    const output = document.getElementById("copilot-prompt-output");
    const hasScenario = !!(appState.currentScenario || (Array.isArray(appState.savedScenarios) && appState.savedScenarios.length));

    if (!hasScenario) {
      if (output) output.value = "";
      if (typeof showToast === "function") showToast("Apply a configuration before generating prompts.", "warning");
      return;
    }

    const ensurePromptGenerated = () => {
      if (typeof window.renderCopilotPromptBundle === "function") {
        try { window.renderCopilotPromptBundle(); } catch (e) {}
      }
      if (output && (!String(output.value || "").trim() || String(output.value || "").includes("Apply a configuration in STEPS"))) {
        // As a last resort, write a minimal prompt instructing the user to regenerate.
        output.value = "Use the Refresh prompt button to regenerate the full STEPS policy prompt.";
      }
    };

    const doCopy = async (afterCopy) => {
      ensurePromptGenerated();
      const text = output ? output.value : "";
      const ok = await copyTextToClipboardRobust(text);
      if (typeof showToast === "function") {
        showToast(ok ? "Prompt copied to clipboard." : "Copy failed. Please copy from the text box.", ok ? "success" : "warning");
      }
      if (typeof afterCopy === "function") afterCopy();
    };

    if (id === "copilot-refresh-btn") {
      ensurePromptGenerated();
      if (typeof showToast === "function") showToast("Prompt refreshed.", "success");
      return;
    }

    if (id === "copilot-download-btn") {
      ensurePromptGenerated();
      const payload = output ? String(output.value || "") : "";
      if (!payload.trim()) {
        if (typeof showToast === "function") showToast("Generate a prompt first.", "warning");
        return;
      }
      const blob = new Blob([payload], { type: "text/plain;charset=utf-8" });
      triggerDownload(blob, `STEPS_prompts_${formatDateForFilename(new Date())}.txt`);
      if (typeof showToast === "function") showToast("Prompts downloaded.", "success");
      return;
    }

    if (id === "copilot-open-and-copy-btn") {
      doCopy(() => window.open("https://copilot.microsoft.com", "_blank", "noopener"));
      return;
    }

    if (id === "chatgpt-open-and-copy-btn") {
      doCopy(() => window.open("https://chat.openai.com", "_blank", "noopener"));
      return;
    }

    // Copy buttons
    doCopy();
  }, true);
}


/* ===========================
   Snapshot modal
   =========================== */

let snapshotModal = null;

function ensureSnapshotModal() {
  if (snapshotModal) return;
  snapshotModal = document.createElement("div");
  snapshotModal.className = "modal hidden";
  snapshotModal.innerHTML = `
    <div class="modal-content">
      <button class="modal-close" type="button" aria-label="Close">×</button>
      <h2>Scenario summary</h2>
      <div id="snapshot-body"></div>
    </div>
  `;
  document.body.appendChild(snapshotModal);

  const closeBtn = snapshotModal.querySelector(".modal-close");
  closeBtn.addEventListener("click", () => {
    snapshotModal.classList.add("hidden");
  });
  snapshotModal.addEventListener("click", (e) => {
    if (e.target === snapshotModal) snapshotModal.classList.add("hidden");
  });
}

function openSnapshotModal(scenario) {
  ensureSnapshotModal();
  const body = snapshotModal.querySelector("#snapshot-body");
  if (body) {
    const c = scenario.config;
    const a = buildAssumptionsForScenario(scenario);

    const mentorBase = scenario.costs?.mentorSupportCostPerCohortBase ?? c.mentorSupportCostPerCohortBase ?? 0;
    const mentorMult = scenario.costs?.mentorCostMultiplier ?? mentorshipMultiplier(c.mentorship);
    const mentorPerCohort = scenario.costs?.mentorCostPerCohort ?? 0;
    const mentorAll = (mentorPerCohort || 0) * (c.cohorts || 0);

    const directPerCohort = scenario.costs?.directCostPerCohort ?? (scenario.costs?.programmeCostPerCohort || 0) + mentorPerCohort;
    const economicPerCohort = scenario.costs?.totalEconomicCostPerCohort ?? 0;

    const cap = scenario.capacity || computeCapacity(c);

    body.innerHTML = `
      <p><strong>Scenario name:</strong> ${safeText(c.name || "")}</p>
      <p><strong>Tier:</strong> ${safeText(c.tier)}</p>
      <p><strong>Career incentive:</strong> ${safeText(c.career)}</p>
      <p><strong>Mentorship:</strong> ${safeText(c.mentorship)}</p>
      <p><strong>Delivery mode:</strong> ${safeText(c.delivery)}</p>
      <p><strong>Response time:</strong> ${safeText(c.response)} days</p>
      <p><strong>Cohorts and trainees:</strong> ${formatNumber(c.cohorts, 0)} cohorts of ${formatNumber(c.traineesPerCohort, 0)} trainees</p>
      <p><strong>Cost per trainee per month:</strong> ${formatCurrencyDisplay(c.costPerTraineePerMonth, 0)}</p>

      <hr />

      <p><strong>Endorsement:</strong> ${formatNumber(scenario.endorseRate, 1)}%</p>
      <p><strong>Perceived programme value per trainee per month:</strong> ${formatCurrencyDisplay(scenario.wtpPerTraineePerMonth, 0)}</p>
      <p><strong>Total perceived programme value (all cohorts):</strong> ${formatCurrencyDisplay(scenario.wtpAllCohorts, 0)}</p>

      <hr />

      <p><strong>Mentor support cost per cohort (base):</strong> ${formatCurrencyDisplay(mentorBase, 0)} (multiplier ${formatNumber(mentorMult, 1)} applied)</p>
      <p><strong>Mentor support cost per cohort:</strong> ${formatCurrencyDisplay(mentorPerCohort, 0)}</p>
      <p><strong>Total mentor support cost (all cohorts):</strong> ${formatCurrencyDisplay(mentorAll, 0)}</p>

      <p><strong>Direct cost per cohort:</strong> ${formatCurrencyDisplay(directPerCohort, 0)}</p>
      <p><strong>Existing opportunity cost per cohort:</strong> ${formatCurrencyDisplay(scenario.costs?.opportunityCostPerCohort ?? 0, 0)} ${c.opportunityCostIncluded ? "" : "(computed as 0 because the switch is off)"}</p>
      <p><strong>Salary-based opportunity cost per cohort (additional):</strong> ${formatCurrencyDisplay(scenario.costs?.salaryBasedOpportunityCostPerCohort ?? 0, 0)} ${c.opportunityCostIncluded ? "" : `(excluded from totals; raw estimate ${formatCurrencyDisplay(scenario.costs?.salaryBasedOpportunityCostPerCohortRaw ?? 0, 0)})`}</p>
      <p><strong>Total economic cost per cohort:</strong> ${formatCurrencyDisplay(economicPerCohort, 0)} ${c.opportunityCostIncluded ? "(includes opportunity cost components)" : "(excludes opportunity cost components)"}</p>

      <p><strong>Total economic cost all cohorts:</strong> ${formatCurrencyDisplay(scenario.natTotalCost, 0)}</p>

      <hr />

      <p><strong>Total indicative epidemiological benefit (per cohort):</strong> ${formatCurrencyDisplay(scenario.epiBenefitPerCohort, 0)}</p>
      <p><strong>Net epidemiological benefit (per cohort):</strong> ${formatCurrencyDisplay(scenario.netBenefitPerCohort, 0)}</p>
      <p><strong>Benefit cost ratio (per cohort):</strong> ${scenario.bcrPerCohort !== null ? formatNumber(scenario.bcrPerCohort, 2) : "-"}</p>

      <p><strong>Total indicative epidemiological benefit (all cohorts):</strong> ${formatCurrencyDisplay(scenario.epiBenefitAllCohorts, 0)}</p>
      <p><strong>Net epidemiological benefit (all cohorts):</strong> ${formatCurrencyDisplay(scenario.netBenefitAllCohorts, 0)}</p>
      <p><strong>National benefit cost ratio:</strong> ${scenario.natBcr !== null ? formatNumber(scenario.natBcr, 2) : "-"}</p>

      <hr />

      <p><strong>Capacity and feasibility:</strong> ${safeText(cap.status)}</p>
      <p><strong>Required mentors per cohort:</strong> ${formatNumber(cap.mentorsPerCohort, 0)} (capacity: ${formatNumber(cap.fellowsPerMentor, 1)} fellows/mentor)</p>
      <p><strong>Total mentors required nationally:</strong> ${formatNumber(cap.totalMentorsRequired, 0)}</p>
      <p><strong>Available mentors nationally:</strong> ${formatNumber(cap.availableMentors, 0)}</p>
      <p><strong>Mentor shortfall:</strong> ${formatNumber(cap.mentorShortfall, 0)}</p>

      <hr />

      <p><strong>Assumptions used:</strong></p>
      <p>Planning horizon: ${formatNumber(a.planningHorizonYears, 0)} years; Discount rate: ${formatNumber(a.discountRate * 100, 1)}%</p>
      <p>Completion rate: ${formatNumber(a.completionRate * 100, 1)}%; Outbreak responses per graduate per year: ${formatNumber(a.outbreaksPerGraduatePerYear, 2)}</p>
      <p>Value per outbreak: ${formatCurrencyDisplay(a.valuePerOutbreak, 0)}; Non-outbreak value per graduate per year: ${formatCurrencyDisplay(a.valuePerGraduate, 0)}</p>
      <p>Cross-sector benefit multiplier: ${formatNumber(a.crossSectorBenefitMultiplier, 2)}</p>
    `;
  }
  snapshotModal.classList.remove("hidden");
}


/* ===========================
   Event wiring and refresh
   =========================== */

function refreshAllOutputs(scenario) {
  updateCostSliderLabel();
  updateConfigSummary(scenario);
  updateResultsTab(scenario);
  updateCostingTab(scenario);
  updateNationalSimulationTab(scenario);
  updateUptakeChart(scenario);
  updateBcrChart(scenario);
  updateEpiChart(scenario);
  refreshSensitivityTables();
  refreshSavedScenariosTable();
  syncOutbreakValueDropdownsFromState();

  // Keep export Enablers / Risks textareas in sync with the current scenario
  const enablersEl = document.getElementById("export-enablers");
  const risksEl = document.getElementById("export-risks");
  if (scenario && scenario.config) {
    if (enablersEl && scenario.config.exportEnablers) {
      enablersEl.value = scenario.config.exportEnablers;
    }
    if (risksEl && scenario.config.exportRisks) {
      risksEl.value = scenario.config.exportRisks;
    }
    updateValidationWarnings(scenario.config);
  } else if (scenario && !scenario.config) {
    updateValidationWarnings({});
  }

  try { if (typeof renderScenariosBaselineCompare === 'function') renderScenariosBaselineCompare(); } catch(e) {}

  try { if (appState && typeof appState.recomputePlanner === 'function') appState.recomputePlanner(); } catch(e) {}

  try { if (typeof window.renderCopilotPromptBundle === 'function') window.renderCopilotPromptBundle(); } catch(e) {}
}
function initEventHandlers() {
  const costSlider = document.getElementById("cost-slider");
  if (costSlider) {
    costSlider.addEventListener("input", () => updateCostSliderLabel());
  }


  // STEPS UPGRADE: enforce cohort capacity constraints live
  const tierEl = document.getElementById("program-tier");
  const cohortsEl = document.getElementById("cohorts");
  const horizonConfigEl = document.getElementById("planning-horizon-config");
  const horizonSettingsEl = document.getElementById("adv-planning-horizon");
  const sitesEl = document.getElementById("available-training-sites");

  const capacityRefresh = (source) => {
    enforceCohortCapacityConstraints({
      tierKey: tierEl ? tierEl.value : (appState.currentScenario && appState.currentScenario.config ? appState.currentScenario.config.tier : "frontline"),
      clampIfNeeded: true,
      source: source || "" 
    });

    // Horizon and sites are treated as operational planning parameters.
    // Update horizon-dependent outputs live so users see immediate changes.
    if (source === "horizon" || source === "sites") {
      scheduleAutoRecomputeScenario(`capacity:${source}`);
    }
  };

  if (tierEl) tierEl.addEventListener("change", () => capacityRefresh("tier"));
  if (cohortsEl) cohortsEl.addEventListener("input", () => capacityRefresh("cohorts"));
  if (horizonConfigEl) horizonConfigEl.addEventListener("input", () => capacityRefresh("horizon"));
  if (horizonSettingsEl) horizonSettingsEl.addEventListener("input", () => {
    // Keep config horizon aligned if user changes Settings
    const v = Number(horizonSettingsEl.value);
    if (isFinite(v) && v > 0 && horizonConfigEl) horizonConfigEl.value = String(v);
    capacityRefresh("horizon");
  });
  if (sitesEl) {
    sitesEl.addEventListener("input", () => {
      // If the user types a value, enforce min 1. If blank, allow blank but treat as 1 in capacity logic.
      const raw = String(sitesEl.value || "").trim();
      if (raw && Number(raw) < 1) {
        sitesEl.value = "1";
      }
      capacityRefresh("sites");
    });
    sitesEl.addEventListener("blur", () => {
      const raw = String(sitesEl.value || "").trim();
      if (raw && Number(raw) < 1) {
        sitesEl.value = "1";
        capacityRefresh("sites");
      }
    });
  }

  const currencyButtons = Array.from(document.querySelectorAll(".pill-toggle"));
  currencyButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const currency = btn.getAttribute("data-currency");
      if (currency && currency !== appState.currency) {
        appState.currency = currency;
        updateCurrencyToggle();
      }
    });
  });

  const oppToggle = document.getElementById("opp-toggle");
  if (oppToggle) {
    oppToggle.addEventListener("click", () => {
      const on = oppToggle.classList.toggle("on");
      const label = oppToggle.querySelector(".switch-label");
      if (label) {
        label.textContent = on ? "Opportunity cost included" : "Opportunity cost excluded";
      }
      if (appState.currentScenario) {
        const newScenario = computeScenario(appState.currentScenario.config);
        appState.currentScenario = newScenario;
        refreshAllOutputs(newScenario);
      }
    });
  }

  const updateBtn = document.getElementById("update-results");
  if (updateBtn) {
    updateBtn.addEventListener("click", () => {
      const config = getConfigFromForm();
      const scenario = computeScenario(config);
      ensureScenarioHasExportNotes(scenario);
      appState.currentScenario = scenario;
      refreshAllOutputs(scenario);
      if (typeof renderBriefBulletsPreview === "function") {
        renderBriefBulletsPreview();
      }
      showToast("Configuration applied and results updated.", "success");
    });
  }

  const snapshotBtn = document.getElementById("open-snapshot");
  if (snapshotBtn) {
    snapshotBtn.addEventListener("click", () => {
      if (!appState.currentScenario) {
        showToast("Apply a configuration before opening the summary.", "warning");
        return;
      }
      openSnapshotModal(appState.currentScenario);
    });
  }

  const saveScenarioBtn = document.getElementById("save-scenario");
  if (saveScenarioBtn) {
    saveScenarioBtn.addEventListener("click", () => {
      if (!appState.currentScenario) {
        showToast("Apply a configuration before saving a scenario.", "warning");
        return;
      }
      appState.savedScenarios.push(cloneScenario(appState.currentScenario));
      refreshSavedScenariosTable();
      refreshSensitivityTables();
      showToast("Scenario saved for comparison and export.", "success");
    });
  }

  const exportExcelBtn = document.getElementById("export-excel");
  if (exportExcelBtn) exportExcelBtn.addEventListener("click", () => exportScenariosToExcel());

  const exportPdfBtn = document.getElementById("export-pdf");
  if (exportPdfBtn) exportPdfBtn.addEventListener("click", () => exportScenariosToPdf());

  const top5RankBy = document.getElementById("top5-rank-by");
  if (top5RankBy) top5RankBy.addEventListener("change", () => refreshTopScenariosPanel("top5"));

  const top5FeasibleOnly = document.getElementById("top5-feasible-only");
  if (top5FeasibleOnly) top5FeasibleOnly.addEventListener("change", () => refreshTopScenariosPanel("top5"));


  const top5SelectedOnly = document.getElementById("top5-selected-only");
  if (top5SelectedOnly) top5SelectedOnly.addEventListener("change", () => refreshTopScenariosPanel("top5"));

  const top5ShortlistAllBtn = document.getElementById("top5-shortlist-all");
  if (top5ShortlistAllBtn) {
    top5ShortlistAllBtn.addEventListener("click", () => {
      const rankByEl = document.getElementById("top5-rank-by");
      const feasibleEl = document.getElementById("top5-feasible-only");
      const selectedEl = document.getElementById("top5-selected-only");
      const metric = rankByEl ? rankByEl.value : "netBenefit";
      const feasibleOnly = feasibleEl ? feasibleEl.checked : false;
      const selectedOnly = selectedEl ? selectedEl.checked : false;
      const items = computeTopScenarioItems(metric, feasibleOnly, selectedOnly, 5);
      if (!items.length) {
        showToast("No scenarios available to shortlist.", "warning");
        return;
      }
      items.forEach((it) => {
        const s = appState.savedScenarios[it.idx];
        if (s) s.shortlisted = true;
      });
      refreshSavedScenariosTable();
      refreshSensitivityTables();
      showToast("Top 5 scenarios shortlisted.", "success");
    });
  }

  const exportTop5PdfBtn = document.getElementById("export-top5-pdf");
  if (exportTop5PdfBtn) exportTop5PdfBtn.addEventListener("click", () => exportTop5OnlyPdf());

  const exportTop5ExcelBtn = document.getElementById("export-top5-excel");
  if (exportTop5ExcelBtn) exportTop5ExcelBtn.addEventListener("click", () => exportTop5OnlyExcel());

  // Copilot tab Top 5 snapshot controls
  const top5CopilotRankBy = document.getElementById("top5-copilot-rank-by");
  const top5CopilotFeasibleOnly = document.getElementById("top5-copilot-feasible-only");
  const top5CopilotSelectedOnly = document.getElementById("top5-copilot-selected-only");
  if (top5CopilotRankBy) top5CopilotRankBy.addEventListener("change", () => refreshTopScenariosPanel("top5-copilot"));
  if (top5CopilotFeasibleOnly) top5CopilotFeasibleOnly.addEventListener("change", () => refreshTopScenariosPanel("top5-copilot"));
  if (top5CopilotSelectedOnly) top5CopilotSelectedOnly.addEventListener("change", () => refreshTopScenariosPanel("top5-copilot"));


  const autoNameBtn = document.getElementById("auto-scenario-name");
  if (autoNameBtn) {
    autoNameBtn.addEventListener("click", () => {
      appState.autoScenarioName = true;
      appState._lastAutoNameSignature = "";
      const cfg = getConfigFromUI(); // updates the input in auto mode
      const scenarioNameEl = document.getElementById("scenario-name");
      const name = (cfg && cfg.name) ? cfg.name : (scenarioNameEl ? (scenarioNameEl.value || "").trim() : "");
      if (scenarioNameEl) scenarioNameEl.value = name;
      if (appState.currentScenario && appState.currentScenario.config) appState.currentScenario.config.name = name;
      showToast("Scenario name refreshed from the current configuration.", "success");
      refreshTopScenariosPanel("top5");
      refreshTopScenariosPanel("top5-copilot");
    });
  }

  const sensUpdateBtn = document.getElementById("refresh-sensitivity-benefits");
  if (sensUpdateBtn) {
    sensUpdateBtn.addEventListener("click", () => {
      if (!appState.currentScenario) {
        showToast("Apply a configuration before updating the sensitivity summary.", "warning");
        return;
      }
      refreshSensitivityTables();
      showToast("Sensitivity summary updated.", "success");
    });
  }

  const sensExcelBtn = document.getElementById("export-sensitivity-benefits-excel");
  if (sensExcelBtn) sensExcelBtn.addEventListener("click", () => exportSensitivityToExcel());

  const epiToggle = getElByIdCandidates(["sensitivity-epi-toggle", "sensitivityEpiToggle"]);
  if (epiToggle) {
    epiToggle.addEventListener("click", () => {
      const on = epiToggle.classList.toggle("on");
      const label = epiToggle.querySelector(".switch-label");
      if (label) label.textContent = on ? "Outbreak benefits included" : "Outbreak benefits excluded";
      if (appState.currentScenario) refreshSensitivityTables();
    });
  }

  const outbreakPresetSelect = getElByIdCandidates(["outbreak-value-preset", "outbreakValuePreset", "outbreak-value"]);
  if (outbreakPresetSelect) {
    ensureSelectHasOutbreakPresets(outbreakPresetSelect);
    outbreakPresetSelect.addEventListener("change", () => {
      const valueInINR = parseSensitivityValueToINR(outbreakPresetSelect.value);
      if (valueInINR) {
        applyOutbreakPreset(valueInINR);
      }
    });
  }

  const outbreakApplyBtn = getElByIdCandidates(["apply-outbreak-value", "applyOutbreakValue", "applyOutbreakPreset"]);
  if (outbreakApplyBtn && outbreakPresetSelect) {
    outbreakApplyBtn.addEventListener("click", () => {
      const valueInINR = parseSensitivityValueToINR(outbreakPresetSelect.value);
      if (valueInINR) {
        applyOutbreakPreset(valueInINR);
      } else {
        showToast("Select a value per outbreak before applying.", "warning");
      }
    });
  }

  const benefitDefSelect = getElByIdCandidates(["benefit-definition-select", "benefitDefinitionSelect"]);
  if (benefitDefSelect) {
    benefitDefSelect.addEventListener("change", () => {
      if (!appState.currentScenario) return;
      refreshSensitivityTables();
    });
  }

  const endorsementOverrideInput = getElByIdCandidates(["endorsement-override", "endorsementOverride"]);
  if (endorsementOverrideInput) {
    endorsementOverrideInput.addEventListener("change", () => {
      if (!appState.currentScenario) return;
      refreshSensitivityTables();
    });
  }

  initApplySettingsButton();
  initSensitivityContractControls();
}




function initCapacityCostLogging() {
  const fields = [
    { id: "mentor-support-cost-per-cohort", label: "Mentor support cost per cohort (base, INR)" },
    { id: "available-mentors-national", label: "Available mentors nationally" },
    { id: "available-training-sites", label: "Available training sites / hubs" },
    { id: "max-cohorts-per-site", label: "Max cohorts per site per year" },
    { id: "cross-sector-multiplier", label: "Cross-sector benefit multiplier" }
  ];

  const opp = document.getElementById("include-opp-cost");
  if (opp) fields.push({ id: "include-opp-cost", label: "Include opportunity cost", type: "checkbox" });

  fields.forEach((f) => {
    const el = document.getElementById(f.id);
    if (!el) return;

    const readVal = () => (f.type === "checkbox" ? !!el.checked : (el.value ?? ""));
    let prev = readVal();

    el.addEventListener("change", () => {
      const now = readVal();
      if (String(now) === String(prev)) return;
      prev = now;
      logSettingsMessage(`${f.label} updated to ${f.type === "checkbox" ? (now ? "ON" : "OFF") : now}.`);
    });
  });
}


/* ===========================
   Initialise
   =========================== */



/* ===========================
   STEPS UPGRADE: Status quo copy to clipboard
   =========================== */

const STATUS_QUO_POLICY_NOTE_TEXT = [
  "Status quo in India (for policy notes)",
  "",
  "Benchmark need: around 1 trained field epidemiologist per 200,000 population (IHR / Global Field Epidemiology Roadmap framing). With population of about 1.3 billion, this implies a long-run need of approximately 6,500 Intermediate/Advanced field epidemiologists.",
  "Supporting point: the CDC India country profile reports India needs over 7,000 epidemiologists to meet IHR targets.",
  "Current stock used in STEPS: approximately 3,300 trained across EIS and FETP tiers (programme estimate), implying a gap of roughly 3,200 Intermediate/Advanced-equivalent graduates needed by 2030. STEPS is designed to test alternative training configurations and scale-up pathways to close this gap.",
  "Existing architecture: India runs a three-tier cascade under NCDC stewardship: Frontline (~3 months), Intermediate (~12 months), and Advanced/EIS (2 years). India's EIS was established in 2012 as a 2-year, mentorship-based in-service programme, complemented by Intermediate and Frontline.",
  "As of 2024: multiple hubs, training a few hundred annually, but below the implied annual flow of about 450-500 Intermediate/Advanced graduates per year needed to close the target by 2030.",
  "Status quo summary: strong architecture and experience, but insufficient volume and uneven distribution, especially at Intermediate and Advanced."
].join("\n");

let statusQuoModal = null;
function ensureStatusQuoModal() {
  if (statusQuoModal) return;
  statusQuoModal = document.createElement("div");
  statusQuoModal.className = "modal hidden";
  statusQuoModal.innerHTML = `
    <div class="modal-content">
      <button class="modal-close" type="button" aria-label="Close">×</button>
      <h2>Copy text for policy note</h2>
      <p class="hint">Select and copy the text below.</p>
      <textarea id="status-quo-textarea" rows="10" style="width:100%;">${STATUS_QUO_POLICY_NOTE_TEXT.replace(/</g, "&lt;")}</textarea>
    </div>
  `;
  document.body.appendChild(statusQuoModal);

  const closeBtn = statusQuoModal.querySelector(".modal-close");
  if (closeBtn) closeBtn.addEventListener("click", () => statusQuoModal.classList.add("hidden"));
  statusQuoModal.addEventListener("click", (e) => {
    if (e.target === statusQuoModal) statusQuoModal.classList.add("hidden");
  });
}

async function copyStatusQuoToClipboard() {
  const statusEl = document.getElementById("copy-status-quo-status");
  const setStatus = (txt) => { if (statusEl) statusEl.textContent = txt || ""; };

  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(STATUS_QUO_POLICY_NOTE_TEXT);
      setStatus("Copied to clipboard.");
      showToast("Status quo text copied.", "success");
      return;
    }
  } catch (e) {
    // Fall through to modal
  }

  ensureStatusQuoModal();
  const ta = statusQuoModal.querySelector("#status-quo-textarea");
  if (ta) {
    ta.value = STATUS_QUO_POLICY_NOTE_TEXT;
    ta.focus();
    ta.select();
  }
  statusQuoModal.classList.remove("hidden");
  setStatus("Clipboard not available. Copy from the pop-up text box.");
}


/* ===========================
   BASELINE + PORTFOLIO + PLANNER (v3)
   =========================== */

// Ensure helpers exist (some older builds may miss these)
if (typeof randomId !== "function") {
  function randomId(prefix = "id") {
    const rnd = Math.random().toString(36).slice(2, 8);
    return `${prefix}_${Date.now().toString(36)}_${rnd}`;
  }
}

if (typeof signedNumber !== "function") {
  function signedNumber(value, decimals = 0) {
    const n = Number(value);
    if (!isFinite(n)) return "-";
    const sign = n > 0 ? "+" : n < 0 ? "-" : "";
    return sign + formatNumber(Math.abs(n), decimals);
  }
}

if (typeof signedCurrency !== "function") {
  function signedCurrency(value, decimals = 0) {
    const n = Number(value);
    if (!isFinite(n)) return "-";
    const sign = n > 0 ? "+" : n < 0 ? "-" : "";
    return sign + formatCurrencyDisplay(Math.abs(n), decimals);
  }
}

// Assumptions helpers used by exports and prompt bundling
function buildAssumptionsBoxData(activeScenario, baselineScenario, deltas) {
  const cfg = (activeScenario && activeScenario.config) ? activeScenario.config : getConfigFromForm();
  const baselineCfg = appState.baselineConfig || buildDefaultBaselineConfig();

  const oppIncluded = !!cfg.opportunityCostIncluded;
  const horizonYears = Number(cfg.planningHorizonYears || getPlanningHorizonFromInputs() || 1);

  const general = {
    planningHorizonYears: horizonYears,
    discountRate: Number(appState.epiSettings?.general?.discountRate ?? 0.03),
    outbreakValueINR: Number(appState.epiSettings?.general?.outbreakValueINR ?? 0),
    outbreakFrequencyPerTraineePerYear: Number(appState.epiSettings?.general?.outbreakFrequencyPerTraineePerYear ?? 0.2),
    benefitCaptureShare: Number(appState.epiSettings?.general?.benefitCaptureShare ?? 0.6),
    opportunityCostIncluded: oppIncluded,
    crossSectorBenefitMultiplier: Number(cfg.crossSectorBenefitMultiplier ?? 1.0)
  };

  function tierSummaryFromConfig(c) {
    if (!c) return null;
    return {
      tier: c.tier,
      cohorts: Number(c.cohorts || 0),
      traineesPerCohort: Number(c.traineesPerCohort || 0),
      costPerTraineePerMonth: Number(c.costPerTraineePerMonth || 0),
      mentorship: c.mentorship,
      delivery: c.delivery,
      career: c.career,
      completionPctOverride: c.completionRateOverride != null ? Number(c.completionRateOverride)*100 : null
    };
  }

  const scenarioInputs = cfg.isPortfolio && cfg.portfolioTiers ? {
    type: "portfolio",
    planningHorizonYears: horizonYears,
    tiers: {
      frontline: tierSummaryFromConfig(cfg.portfolioTiers.frontline),
      intermediate: tierSummaryFromConfig(cfg.portfolioTiers.intermediate),
      advanced: tierSummaryFromConfig(cfg.portfolioTiers.advanced)
    },
    availableTrainingSites: Number(cfg.availableTrainingSites ?? baselineCfg.availableTrainingSites ?? 1),
    availableMentorsNational: Number(cfg.availableMentorsNational ?? 0),
    mentorSupportCostPerCohortBase: Number(cfg.mentorSupportCostPerCohortBase ?? 0)
  } : {
    type: "single",
    ...tierSummaryFromConfig(cfg),
    planningHorizonYears: horizonYears,
    availableTrainingSites: Number(cfg.availableTrainingSites ?? 1),
    availableMentorsNational: Number(cfg.availableMentorsNational ?? 0),
    mentorSupportCostPerCohortBase: Number(cfg.mentorSupportCostPerCohortBase ?? 0)
  };

  const baselineInputs = {
    availableTrainingSites: Number(baselineCfg.availableTrainingSites || 1),
    tiers: baselineCfg.tiers
  };

  const outputs = {
    baseline: baselineScenario ? {
      totalGraduates: baselineScenario.graduatesAllCohorts,
      iaGraduates: getIntermediateAdvancedGraduates(baselineScenario),
      totalCost: baselineScenario.natTotalCost,
      totalBenefit: baselineScenario.epiBenefitAllCohorts,
      netBenefit: baselineScenario.netBenefitAllCohorts,
      bcr: baselineScenario.natBcr
    } : null,
    scenario: activeScenario ? {
      totalGraduates: activeScenario.graduatesAllCohorts,
      iaGraduates: getIntermediateAdvancedGraduates(activeScenario),
      totalCost: activeScenario.natTotalCost,
      totalBenefit: activeScenario.epiBenefitAllCohorts,
      netBenefit: activeScenario.netBenefitAllCohorts,
      bcr: activeScenario.natBcr
    } : null,
    incremental: deltas || null
  };

  return { general, baselineInputs, scenarioInputs, outputs };
}

function formatAssumptionsBoxText(data) {
  if (!data) return "";
  const g = data.general || {};
  const lines = [];
  lines.push("STEPS assumptions snapshot");
  lines.push(`Planning horizon: ${formatNumber(g.planningHorizonYears || 0, 0)} year(s)`);
  lines.push(`Discount rate: ${formatNumber((g.discountRate || 0) * 100, 1)}%`);
  lines.push(`Outbreak value: ${formatCurrencyDisplay(g.outbreakValueINR || 0, 0)}`);
  lines.push(`Outbreak frequency (per trainee-year): ${formatNumber(g.outbreakFrequencyPerTraineePerYear || 0, 2)}`);
  lines.push(`Benefit capture share: ${formatNumber((g.benefitCaptureShare || 0) * 100, 0)}%`);
  lines.push(`Opportunity cost included: ${g.opportunityCostIncluded ? "Yes" : "No"}`);
  lines.push(`Cross-sector benefit multiplier: ${formatNumber(g.crossSectorBenefitMultiplier || 1, 2)}x`);

  function tierBlock(title, t) {
    if (!t) return;
    lines.push("");
    lines.push(title);
    if (t.type === "portfolio") {
      const tiers = t.tiers || {};
      for (const key of ["frontline","intermediate","advanced"]) {
        const z = tiers[key];
        if (!z) continue;
        lines.push(`  ${key}: cohorts=${formatNumber(z.cohorts||0,0)}, trainees/cohort=${formatNumber(z.traineesPerCohort||0,0)}, cost/trainee-month=${formatCurrencyDisplay(z.costPerTraineePerMonth||0,0)}${z.completionPctOverride!=null?`, completion=${formatNumber(z.completionPctOverride,0)}%`:""}`);
      }
      lines.push(`  Available training sites/hubs: ${formatNumber(t.availableTrainingSites||1,0)}`);
    } else {
      lines.push(`  Tier: ${safeText(t.tier || "")}`);
      lines.push(`  Cohorts: ${formatNumber(t.cohorts||0,0)}; trainees per cohort: ${formatNumber(t.traineesPerCohort||0,0)}`);
      lines.push(`  Cost per trainee-month: ${formatCurrencyDisplay(t.costPerTraineePerMonth||0,0)}`);
      if (t.completionPctOverride != null) lines.push(`  Completion override: ${formatNumber(t.completionPctOverride,0)}%`);
      lines.push(`  Available training sites/hubs: ${formatNumber(t.availableTrainingSites||1,0)}`);
    }
  }

  tierBlock("Baseline inputs (BAU)", { type: "portfolio", tiers: data.baselineInputs?.tiers || {}, availableTrainingSites: data.baselineInputs?.availableTrainingSites || 1 });
  tierBlock("Scenario inputs", data.scenarioInputs);

  const out = data.outputs || {};
  function outBlock(title, o) {
    if (!o) return;
    lines.push("");
    lines.push(title);
    lines.push(`  Total graduates: ${formatNumber(o.totalGraduates||0,0)}`);
    lines.push(`  Intermediate+Advanced graduates: ${formatNumber(o.iaGraduates||0,0)}`);
    lines.push(`  Total cost: ${formatCurrencyDisplay(o.totalCost||0,0)}`);
    lines.push(`  Total benefit: ${formatCurrencyDisplay(o.totalBenefit||0,0)}`);
    lines.push(`  Net benefit: ${formatCurrencyDisplay(o.netBenefit||0,0)}`);
    lines.push(`  BCR: ${o.bcr!=null?formatNumber(o.bcr,2):"-"}`);
  }
  outBlock("Baseline outputs", out.baseline);
  outBlock("Scenario outputs", out.scenario);
  outBlock("Incremental (Scenario minus Baseline)", out.incremental);

  return lines.join("\n");
}

function getIntermediateAdvancedGraduates(scenario) {
  if (!scenario) return 0;
  const cfg = scenario.config || {};
  if (cfg && cfg.isPortfolio && cfg.portfolioTiers) {
    const br = scenario.portfolioBreakdown || scenario._portfolioBreakdown || {};
    const i = br.intermediate?.graduatesAllCohorts ?? 0;
    const a = br.advanced?.graduatesAllCohorts ?? 0;
    return Number(i) + Number(a);
  }
  if (cfg.tier === "intermediate" || cfg.tier === "advanced") return Number(scenario.graduatesAllCohorts || 0);
  return 0;
}



function getFrontlineGraduates(scenario) {
  if (!scenario) return 0;
  const cfg = scenario.config || {};
  if (cfg && cfg.isPortfolio && cfg.portfolioTiers) {
    const br = scenario.portfolioBreakdown || scenario._portfolioBreakdown || {};
    const f = br.frontline?.graduatesAllCohorts ?? 0;
    return Number(f) || 0;
  }
  if (cfg.tier === "frontline") return Number(scenario.graduatesAllCohorts || 0);
  return 0;
}

function getIntermediateGraduates(scenario) {
  if (!scenario) return 0;
  const cfg = scenario.config || {};
  if (cfg && cfg.isPortfolio && cfg.portfolioTiers) {
    const br = scenario.portfolioBreakdown || scenario._portfolioBreakdown || {};
    const i = br.intermediate?.graduatesAllCohorts ?? 0;
    return Number(i) || 0;
  }
  if (cfg.tier === "intermediate") return Number(scenario.graduatesAllCohorts || 0);
  return 0;
}

function getAdvancedGraduates(scenario) {
  if (!scenario) return 0;
  const cfg = scenario.config || {};
  if (cfg && cfg.isPortfolio && cfg.portfolioTiers) {
    const br = scenario.portfolioBreakdown || scenario._portfolioBreakdown || {};
    const a = br.advanced?.graduatesAllCohorts ?? 0;
    return Number(a) || 0;
  }
  if (cfg.tier === "advanced") return Number(scenario.graduatesAllCohorts || 0);
  return 0;
}

function computePortfolioScenario(portfolioConfig) {
  const cfg = portfolioConfig || {};
  const tiers = cfg.portfolioTiers || {};

  const tierKeys = ["frontline","intermediate","advanced"];
  const breakdown = {};

  let totalCost = 0;
  let totalBenefit = 0;
  let totalGrads = 0;
  let outbreaksPerYear = 0;

  let endorseWeighted = 0;
  let optOutWeighted = 0;
  let weightDenom = 0;

  let wtpAll = 0;

  for (const k of tierKeys) {
    const tc = tiers[k];
    if (!tc) continue;
    const s = computeScenario(tc);
    breakdown[k] = s;

    const w = Number(tc.cohorts || 0) * Number(tc.traineesPerCohort || 0);
    endorseWeighted += (s.endorseRate || 0) * w;
    optOutWeighted += (s.optOutRate || 0) * w;
    weightDenom += w;

    totalCost += Number(s.natTotalCost || 0);
    totalBenefit += Number(s.epiBenefitAllCohorts || 0);
    totalGrads += Number(s.graduatesAllCohorts || 0);
    outbreaksPerYear += Number(s.outbreaksPerYearNational || 0);
    wtpAll += Number(s.wtpAllCohorts || 0);
  }

  const net = totalBenefit - totalCost;
  const bcr = totalCost > 0 ? (totalBenefit / totalCost) : null;

  const endorseRate = weightDenom > 0 ? (endorseWeighted / weightDenom) : 0;
  const optOutRate = weightDenom > 0 ? (optOutWeighted / weightDenom) : Math.max(0, 100 - endorseRate);

  return {
    id: cfg.id || randomId("pf"),
    _sid: cfg._sid || randomId("sc"),
    timestamp: new Date().toISOString(),
    shortlisted: false,
    pinned: !!cfg.pinned,
    pinnedAt: cfg.pinnedAt || null,
    createdAt: cfg.createdAt || Date.now(),
    config: cfg,
    preferenceModel: cfg.preferenceModel || "Mixed logit model from the preference study",
    endorseRate,
    optOutRate,

    // Aggregate outputs used throughout the UI
    wtpAllCohorts: wtpAll,
    wtpPerTraineePerMonth: weightDenom > 0 ? (wtpAll / (weightDenom * Math.max(1, Number(cfg.planningHorizonYears||1)) * 12)) : 0,
    wtpPerCohort: null,
    wtpOutbreakComponent: wtpAll * 0.3,

    costs: {
      totalEconomicCostPerCohort: null,
      totalEconomicCostAllCohorts: totalCost
    },
    epi: {
      graduatesAllCohorts: totalGrads,
      outbreaksPerYearNational: outbreaksPerYear
    },
    capacity: computeCapacity(cfg.portfolioTiers?.advanced || cfg.portfolioTiers?.intermediate || cfg.portfolioTiers?.frontline || cfg),

    epiBenefitAllCohorts: totalBenefit,
    netBenefitAllCohorts: net,
    natTotalCost: totalCost,
    natBcr: bcr,
    graduatesAllCohorts: totalGrads,
    outbreaksPerYearNational: outbreaksPerYear,
    planningYears: Number(cfg.planningHorizonYears || 1),

    // keep breakdown for IA calculations and detailed cards
    portfolioBreakdown: breakdown
  };
}

function computeBaselineAggregate(horizonYears) {
  const baselineCfg = loadBaselineConfig();
  const baseSites = Math.max(1, Number(baselineCfg.availableTrainingSites || 1));

  const baseScenarioCfg = getConfigFromForm();

  function tierConfigFromBaseline(tierKey) {
    const t = baselineCfg.tiers?.[tierKey] || {};
    const completionPct = Number(t.completionPct);
    const completionRateOverride = isFinite(completionPct) ? clamp(completionPct / 100, 0, 1) : null;
    const career = t.career || baseScenarioCfg.career;
    const mentorship = t.mentorship || baseScenarioCfg.mentorship;
    const delivery = t.delivery || baseScenarioCfg.delivery;

    return {
      tier: tierKey,
      career,
      mentorship,
      delivery,
      response: baseScenarioCfg.response,

      costPerTraineePerMonth: Number(t.costPerTraineePerMonth || 0),
      traineesPerCohort: Math.max(1, Number(t.traineesPerCohort || 1)),
      cohorts: Math.max(0, Number(t.cohorts || 0)),
      planningHorizonYears: Number(horizonYears || baseScenarioCfg.planningHorizonYears || 1),
      opportunityCostIncluded: !!baseScenarioCfg.opportunityCostIncluded,

      mentorSupportCostPerCohortBase: Number(baseScenarioCfg.mentorSupportCostPerCohortBase || 0),
      availableMentorsNational: Number(baseScenarioCfg.availableMentorsNational || 0),
      availableTrainingSites: baseSites,
      maxCohortsPerSitePerYear: Number(baseScenarioCfg.maxCohortsPerSitePerYear || 0),
      crossSectorBenefitMultiplier: Number(baseScenarioCfg.crossSectorBenefitMultiplier || 1.0),

      completionRateOverride: completionRateOverride,

      name: `Baseline ${tierKey}`
    };
  }

  const portfolioConfig = {
    isPortfolio: true,
    tier: "portfolio",
    planningHorizonYears: Number(horizonYears || baseScenarioCfg.planningHorizonYears || 1),
    availableTrainingSites: baseSites,
    preferenceModel: baseScenarioCfg.preferenceModel,
    portfolioTiers: {
      frontline: tierConfigFromBaseline("frontline"),
      intermediate: tierConfigFromBaseline("intermediate"),
      advanced: tierConfigFromBaseline("advanced")
    },
    name: "Baseline (BAU)",
    notes: "Business as usual / no-change counterfactual"
  };

  const baselineScenario = computePortfolioScenario(portfolioConfig);
  baselineScenario.config = portfolioConfig;
  baselineScenario._portfolioBreakdown = baselineScenario.portfolioBreakdown;
  
  // Attach national planning baseline stocks (Intermediate/Advanced) from Planner tab.
  // These are used in National Simulation baseline–scenario comparisons and Copilot prompts.
  const targetEl = document.getElementById("planner-target-stock");
  const currentEl = document.getElementById("planner-current-stock");
  const targetIA = Math.max(0, Math.round(Number(targetEl ? targetEl.value : 7000) || 0));
  const currentIA = Math.max(0, Math.round(Number(currentEl ? currentEl.value : 3300) || 0));
  baselineScenario.targetStockIA = targetIA;
  baselineScenario.currentStockIA = currentIA;
  baselineScenario.gapToTargetIA = Math.max(0, targetIA - currentIA);

  return baselineScenario;
}


/**
 * Build a national (portfolio) scenario for the National Simulation + Planner gap logic.
 * This keeps the baseline (Planner tab) values for tiers that are NOT currently selected in the Configuration tab,
 * and overrides ONLY the currently configured tier using the Configuration tab values.
 *
 * Rationale: users iterate scenarios via the Configuration tab, while the Baseline editor remains a BAU counterfactual.
 */
function computeScenarioAggregateFromCurrentConfig(horizonYears) {
  const baselineCfg = loadBaselineConfig();
  const baseSites = Math.max(1, Number((baselineCfg.availableTrainingSites ?? baselineCfg.availableSites) || 1));

  const currentCfg = getConfigFromForm();

  // Global operational/assumption parameters are taken from the current configuration (to keep assumptions consistent),
  // while tier-specific BAU scale is taken from the baseline editor.
  const sites = Math.max(1, Number(currentCfg.availableTrainingSites || baseSites));

  function tierConfigFromBaseline(tierKey) {
    const t = baselineCfg.tiers?.[tierKey] || {};
    const completionPct = Number(t.completionPct);
    const completionRateOverride = isFinite(completionPct) ? clamp(completionPct / 100, 0, 1) : null;

    return {
      tier: tierKey,
      career: t.career || currentCfg.career,
      mentorship: t.mentorship || currentCfg.mentorship,
      delivery: t.delivery || currentCfg.delivery,
      response: currentCfg.response,

      costPerTraineePerMonth: Number(t.costPerTraineePerMonth || 0),
      traineesPerCohort: Math.max(0, Number(t.traineesPerCohort || 0)),
      cohorts: Math.max(0, Number(t.cohorts || 0)),
      planningHorizonYears: Number(horizonYears || currentCfg.planningHorizonYears || 1),
      opportunityCostIncluded: !!currentCfg.opportunityCostIncluded,

      mentorSupportCostPerCohortBase: Number(currentCfg.mentorSupportCostPerCohortBase || 0),
      availableMentorsNational: Number(currentCfg.availableMentorsNational || 0),
      availableTrainingSites: sites,
      maxCohortsPerSitePerYear: Number(currentCfg.maxCohortsPerSitePerYear || 0),
      crossSectorBenefitMultiplier: Number(currentCfg.crossSectorBenefitMultiplier || 1.0),

      completionRateOverride: completionRateOverride,
      name: `Scenario aggregate ${tierKey}`
    };
  }

  const selectedTier = (currentCfg.tier || "frontline").toLowerCase();
  const portfolioConfig = {
    isPortfolio: true,
    tier: "portfolio",
    planningHorizonYears: Number(horizonYears || currentCfg.planningHorizonYears || 1),
    availableTrainingSites: sites,
    preferenceModel: currentCfg.preferenceModel,
    portfolioTiers: {
      frontline: tierConfigFromBaseline("frontline"),
      intermediate: tierConfigFromBaseline("intermediate"),
      advanced: tierConfigFromBaseline("advanced")
    },
    name: currentCfg.name || "Scenario aggregate (current configuration)",
    notes: "National aggregate for comparisons: baseline tiers + current configured tier override"
  };

  // Override ONLY the selected tier with the current configuration values
  if (portfolioConfig.portfolioTiers[selectedTier]) {
    portfolioConfig.portfolioTiers[selectedTier] = {
      ...portfolioConfig.portfolioTiers[selectedTier],
      tier: selectedTier,
      career: currentCfg.career,
      mentorship: currentCfg.mentorship,
      delivery: currentCfg.delivery,
      response: currentCfg.response,
      cohorts: Math.max(0, Number(currentCfg.cohorts || 0)),
      traineesPerCohort: Math.max(0, Number(currentCfg.traineesPerCohort || 0)),
      costPerTraineePerMonth: Math.max(0, Number(currentCfg.costPerTraineePerMonth || 0)),
      planningHorizonYears: Number(horizonYears || currentCfg.planningHorizonYears || 1),
      opportunityCostIncluded: !!currentCfg.opportunityCostIncluded,
      availableTrainingSites: sites,
      name: `Scenario ${selectedTier} override`
    };
  }

  const aggScenario = computePortfolioScenario(portfolioConfig);
  aggScenario.config = portfolioConfig;
  aggScenario._portfolioBreakdown = aggScenario.portfolioBreakdown;

  // Convenience: expose selected tier on the returned object for downstream UI logic
  aggScenario._selectedTier = selectedTier;
  return aggScenario;
}


function updateBaselineTierEndorsementsDisplay(baselineScenario) {
  const base = baselineScenario;
  const breakdown = base?.portfolioBreakdown || base?._portfolioBreakdown;
  if (!breakdown) return;

  const mapping = {
    frontline: "baseline-frontline-endorsement",
    intermediate: "baseline-intermediate-endorsement",
    advanced: "baseline-advanced-endorsement"
  };

  Object.keys(mapping).forEach((k) => {
    const el = document.getElementById(mapping[k]);
    if (!el) return;
    const s = breakdown[k];
    const v = s && isFinite(Number(s.endorseRate)) ? Number(s.endorseRate) : null;
    el.textContent = v !== null ? `${formatNumber(v, 1)}%` : "-";
  });
}

function computeIncrementalMetrics(scenario, baseline) {
  if (!scenario || !baseline) return null;
  const dCost = Number(scenario.natTotalCost || 0) - Number(baseline.natTotalCost || 0);
  const dBenefit = Number(scenario.epiBenefitAllCohorts || 0) - Number(baseline.epiBenefitAllCohorts || 0);
  const dNet = dBenefit - dCost;
  let incBcr = null;
  if (dCost > 0) incBcr = dBenefit / dCost;
  return {
    deltaCost: dCost,
    deltaBenefit: dBenefit,
    deltaNet: dNet,
    incrementalBcr: incBcr,
    note: dCost <= 0 ? "Incremental BCR not defined when incremental cost ≤ 0" : null
  };
}

function getBaselineForCurrentHorizon() {
  const horizon = getPlanningHorizonFromInputs();
  const baseline = computeBaselineAggregate(horizon);
  appState.baselineScenario = baseline;
  appState.baselineConfig = loadBaselineConfig();
  return baseline;
}

function setTextIfExists(id, text) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = text;
}



function aggregateScenarioSet(items) {
  const out = {
    name: `Selected package (${items.length} scenarios)`,
    count: items.length
  };
  // Graduates
  const ia = items.reduce((a,s)=>a+Number((s.intermediateGraduatesAllCohorts ?? s.natIntermediateGrads ?? 0) + (s.advancedGraduatesAllCohorts ?? s.natAdvancedGrads ?? 0)),0);
  const inter = items.reduce((a,s)=>a+Number(s.intermediateGraduatesAllCohorts ?? s.natIntermediateGrads ?? 0),0);
  const adv = items.reduce((a,s)=>a+Number(s.advancedGraduatesAllCohorts ?? s.natAdvancedGrads ?? 0),0);
  const front = items.reduce((a,s)=>a+Number(s.frontlineGraduatesAllCohorts ?? s.natFrontlineGrads ?? 0),0);
  const total = items.reduce((a,s)=>a+Number(s.totalGraduatesAllTiersAllCohorts ?? s.natTotalGrads ?? 0),0);

  out.iaGraduates = ia;
  out.intermediateGraduates = inter;
  out.advancedGraduates = adv;
  out.frontlineGraduates = front;
  out.totalGraduates = total;

  // Costs/benefits
  const cost = items.reduce((a,s)=>a+Number(s.totalEconomicCostAllCohorts ?? s.natTotalCost ?? 0),0);
  const benefit = items.reduce((a,s)=>a+Number(s.totalEpiBenefitsAllCohorts ?? s.epiBenefitAllCohorts ?? 0),0);
  out.totalEconomicCost = cost;
  out.totalEpiBenefit = benefit;
  out.netBenefit = benefit - cost;
  out.bcr = cost > 0 ? benefit / cost : 0;

  // Feasibility: if any is not feasible, mark as requires expansion.
  const feasStatuses = items.map((s)=> (s.capacity?.status || (s.config ? computeCapacity(s.config).status : "Unknown")) );
  out.feasibility = feasStatuses.includes("Feasible") && !feasStatuses.includes("Not feasible") ? "Feasible (all selected scenarios)" :
    (feasStatuses.includes("Not feasible") ? "Requires expansion (one or more scenarios not feasible)" : "Mixed / check capacity");
  return out;
}

function buildPackageComparisonTableText(baseline, pkg) {
  if (!pkg) return "";
  const targetIA = Number(baseline.targetStockIA || 0);
  const currentIA = Number(baseline.currentStockIA || 0);
  const baseRemaining = Math.max(0, targetIA - currentIA);
  // Per user requirement: treat total graduates (including Frontline) as contributing to closing the target gap.
  const scenRemaining = Math.max(0, targetIA - (currentIA + Number(pkg.totalGraduates || 0)));

  const rows = [
    ["Indicator", "Baseline", "Scenario package", "Incremental (Scenario − Baseline)"],
    ["Total FETP graduates (all tiers)", formatNumber(Number(baseline.natTotalGrads || 0),0), formatNumber(Number(pkg.totalGraduates||0),0), formatNumber(Number(pkg.totalGraduates||0)-Number(baseline.natTotalGrads||0),0)],
    ["Frontline graduates", formatNumber(Number(baseline.natFrontlineGrads||0),0), formatNumber(Number(pkg.frontlineGraduates||0),0), formatNumber(Number(pkg.frontlineGraduates||0)-Number(baseline.natFrontlineGrads||0),0)],
    ["Intermediate graduates", formatNumber(Number(baseline.natIntermediateGrads||0),0), formatNumber(Number(pkg.intermediateGraduates||0),0), formatNumber(Number(pkg.intermediateGraduates||0)-Number(baseline.natIntermediateGrads||0),0)],
    ["Advanced graduates", formatNumber(Number(baseline.natAdvancedGrads||0),0), formatNumber(Number(pkg.advancedGraduates||0),0), formatNumber(Number(pkg.advancedGraduates||0)-Number(baseline.natAdvancedGrads||0),0)],
    ["Intermediate + Advanced graduates", formatNumber(Number(baseline.natIAGrads||0),0), formatNumber(Number(pkg.iaGraduates||0),0), formatNumber(Number(pkg.iaGraduates||0)-Number(baseline.natIAGrads||0),0)],
    ["Remaining to reach target (all tiers)", formatNumber(baseRemaining,0), formatNumber(scenRemaining,0), formatNumber(scenRemaining-baseRemaining,0)],
    ["Total economic cost (INR)", formatCurrencyDisplay(Number(baseline.natTotalCost||0),0), formatCurrencyDisplay(Number(pkg.totalEconomicCost||0),0), formatCurrencyDisplay(Number(pkg.totalEconomicCost||0)-Number(baseline.natTotalCost||0),0)],
    ["Epidemiological benefit (INR)", formatCurrencyDisplay(Number(baseline.epiBenefitAllCohorts||0),0), formatCurrencyDisplay(Number(pkg.totalEpiBenefit||0),0), formatCurrencyDisplay(Number(pkg.totalEpiBenefit||0)-Number(baseline.epiBenefitAllCohorts||0),0)],
    ["Net benefit (INR)", formatCurrencyDisplay(Number(baseline.netBenefitAllCohorts||0),0), formatCurrencyDisplay(Number(pkg.netBenefit||0),0), formatCurrencyDisplay(Number(pkg.netBenefit||0)-Number(baseline.netBenefitAllCohorts||0),0)],
    ["Benefit–cost ratio (BCR)", formatNumber((Number(baseline.natBcr)||safeBcr(baseline)),2), formatNumber(Number(pkg.bcr||0),2), formatNumber(Number(pkg.bcr||0)-Number((Number(baseline.natBcr)||safeBcr(baseline))),2)]
  ];
  return rows.map(r=>r.join(" | ")).join("\n");
}

function buildScenarioEvidenceTableText(items) {
  const rows = [];
  rows.push("Scenario | IA graduates | Total graduates | Economic cost (INR) | Epidemiological benefit (INR) | Net benefit (INR) | BCR | Feasibility");
  for (const s of items) {
    const name = safeText(getScenarioDisplayName(s));
    const inter = Number(s.intermediateGraduatesAllCohorts ?? s.natIntermediateGrads ?? 0);
    const adv = Number(s.advancedGraduatesAllCohorts ?? s.natAdvancedGrads ?? 0);
    const ia = inter + adv;
    const total = Number(s.totalGraduatesAllTiersAllCohorts ?? s.natTotalGrads ?? 0);
    const cost = Number(s.totalEconomicCostAllCohorts ?? s.natTotalCost ?? 0);
    const ben = Number(s.totalEpiBenefitsAllCohorts ?? s.epiBenefitAllCohorts ?? 0);
    const net = ben - cost;
    const bcr = cost > 0 ? ben / cost : 0;
    const feas = s.capacity ? s.capacity.status : (s.config ? computeCapacity(s.config).status : "Unknown");
    rows.push([name, formatNumber(ia,0), formatNumber(total,0), formatCurrencyDisplay(cost,0), formatCurrencyDisplay(ben,0), formatCurrencyDisplay(net,0), formatNumber(bcr,2), safeText(feas)].join(" | "));
  }
  return rows.join("\n");
}
function renderNationalBaselineScenarioDelta(activeScenario) {
  // If no scenario is passed, use the selection from the baseline comparison selector.
  const baseline = getBaselineForCurrentHorizon();

  const saved = Array.isArray(appState.savedScenarios) ? appState.savedScenarios : [];
  const shortlist = Array.isArray(appState.baselineCompareShortlist) ? appState.baselineCompareShortlist : [];

  // Populate selector options (saved scenarios plus a "current configuration" option).
  const sel = document.getElementById("baseline-scenario-select");
  if (sel) {
    const prev = sel.value;
    sel.innerHTML = "";
    const optCurrent = document.createElement("option");
    optCurrent.value = "__current__";
    optCurrent.textContent = "Current configuration (not saved)";
    sel.appendChild(optCurrent);

    saved.forEach((sc, idx) => {
      const o = document.createElement("option");
      o.value = String(idx);
      o.textContent = getScenarioDisplayName(sc);
      sel.appendChild(o);
    });

    // Determine active selection (shortlist is used for aggregation; dropdown remains a quick selector).
    let activeId = (appState.baselineCompareActive != null) ? String(appState.baselineCompareActive) : null;

    if (activeId && activeId !== "__current__") {
      const ai = Number(activeId);
      if (isNaN(ai) || !saved[ai]) activeId = null;
    }
    if (!activeId && prev && Array.from(sel.options).some((o) => o.value === prev)) activeId = prev;
    if (!activeId && saved.length > 0) activeId = "0";
    if (!activeId) activeId = "__current__";

    sel.value = activeId;

    if (!sel._boundChange) {
      sel.addEventListener("change", () => {
        appState.baselineCompareActive = sel.value;
        try { localStorage.setItem("steps_baseline_compare_active", String(sel.value)); } catch (e) {}
        renderNationalBaselineScenarioDelta(null);
      });
      sel._boundChange = true;
    }

    // Shortlist controls (add selected scenarios as quick-access chips)
    const addBtn = document.getElementById("baseline-scenario-add");
    const chipWrap = document.getElementById("baseline-scenario-chips");

    // Load persisted shortlist on first run
    if (!appState._baselineCompareLoaded) {
      try {
        const raw = localStorage.getItem("steps_baseline_compare_shortlist");
        if (raw) appState.baselineCompareShortlist = JSON.parse(raw) || [];
      } catch (e) { appState.baselineCompareShortlist = []; }

      try {
        const rawA = localStorage.getItem("steps_baseline_compare_active");
        if (rawA != null) appState.baselineCompareActive = rawA;
      } catch (e) { /* ignore */ }

      appState._baselineCompareLoaded = true;
    }

    function persistShortlist() {
      try { localStorage.setItem("steps_baseline_compare_shortlist", JSON.stringify(appState.baselineCompareShortlist || [])); } catch (e) {}
      try { localStorage.setItem("steps_baseline_compare_active", String(appState.baselineCompareActive ?? "__current__")); } catch (e) {}
    }

    function renderChips() {
      if (!chipWrap) return;
      chipWrap.innerHTML = "";
      const list = Array.isArray(appState.baselineCompareShortlist) ? appState.baselineCompareShortlist : [];
      list.forEach((idxStr) => {
        const i = Number(idxStr);
        const sc = saved[i];
        if (!sc) return;
        const chip = document.createElement("button");
        chip.type = "button";
        chip.className = "chip" + (String(appState.baselineCompareActive) === String(idxStr) ? " active" : "");
        chip.setAttribute("data-chip-idx", String(idxStr));
        chip.textContent = getScenarioDisplayName(sc);

        const x = document.createElement("span");
        x.className = "chip-x";
        x.setAttribute("data-chip-remove", String(idxStr));
        x.textContent = "×";
        chip.appendChild(x);

        chipWrap.appendChild(chip);
      });
    }

    if (addBtn && !addBtn._boundClick) {
      addBtn.addEventListener("click", () => {
        if (!sel) return;
        if (sel.value === "__current__") {
          showToast("Please select a saved scenario to add to the comparison list.", "warning");
          return;
        }
        const list = Array.isArray(appState.baselineCompareShortlist) ? appState.baselineCompareShortlist : [];
        if (!list.includes(sel.value)) list.push(sel.value);
        appState.baselineCompareShortlist = list;
        appState.baselineCompareActive = sel.value;
        persistShortlist();
        renderChips();
        renderNationalBaselineScenarioDelta(null);
        showToast("Scenario added to the comparison list.", "success");
      });
      addBtn._boundClick = true;
    }

    if (chipWrap && !chipWrap._boundClick) {
      chipWrap.addEventListener("click", (ev) => {
        const t = ev.target;
        if (!(t instanceof HTMLElement)) return;

        const rem = t.getAttribute("data-chip-remove");
        if (rem != null) {
          ev.preventDefault();
          ev.stopPropagation();
          const list = Array.isArray(appState.baselineCompareShortlist) ? appState.baselineCompareShortlist : [];
          appState.baselineCompareShortlist = list.filter((x) => String(x) !== String(rem));
          if (String(appState.baselineCompareActive) === String(rem)) {
            appState.baselineCompareActive = (appState.baselineCompareShortlist && appState.baselineCompareShortlist.length) ? appState.baselineCompareShortlist[0] : "__current__";
            sel.value = String(appState.baselineCompareActive);
          }
          persistShortlist();
          renderChips();
          renderNationalBaselineScenarioDelta(null);
          showToast("Scenario removed from the comparison list.", "success");
          return;
        }

        const chipIdx = t.getAttribute("data-chip-idx") || t.closest("[data-chip-idx]")?.getAttribute("data-chip-idx");
        if (chipIdx != null) {
          appState.baselineCompareActive = chipIdx;
          sel.value = String(chipIdx);
          persistShortlist();
          renderChips();
          renderNationalBaselineScenarioDelta(null);
        }
      });
      chipWrap._boundClick = true;
    }

    renderChips();
  }

  function aggregateShortlistedScenario(idxList) {
    const list = Array.isArray(idxList) ? idxList : [];
    const seen = new Set();
    const indices = [];
    list.forEach((x) => {
      const i = Number(x);
      if (!isFinite(i) || i < 0) return;
      if (!saved[i]) return;
      const key = String(i);
      if (seen.has(key)) return;
      seen.add(key);
      indices.push(i);
    });
    if (!indices.length) return null;

    let sumFrontline = 0, sumIntermediate = 0, sumAdvanced = 0;
    let sumGrads = 0, sumCost = 0, sumBenefit = 0, sumNet = 0;

    indices.forEach((i) => {
      const s = saved[i];
      if (!s) return;
      sumFrontline += Number(getFrontlineGraduates(s) || 0);
      sumIntermediate += Number(getIntermediateGraduates(s) || 0);
      sumAdvanced += Number(getAdvancedGraduates(s) || 0);
      sumGrads += Number(s.graduatesAllCohorts || 0);

      const sc = Number(s.natTotalCost || s.totalEconomicCostAllCohorts || 0);
      const sb = Number(s.epiBenefitAllCohorts || s.totalEpiBenefitsAllCohorts || 0);
      const sn = (s.netBenefitAllCohorts != null) ? Number(s.netBenefitAllCohorts) : (sb - sc);
      sumCost += sc;
      sumBenefit += sb;
      sumNet += sn;
    });

    const ia = sumIntermediate + sumAdvanced;

    return {
      id: `aggregate_${indices.join("_")}`,
      config: {
        isPortfolio: true,
        portfolioTiers: { frontline: true, intermediate: true, advanced: true },
        name: `Selected scenarios (aggregate, n=${indices.length})`
      },
      portfolioBreakdown: {
        frontline: { graduatesAllCohorts: sumFrontline },
        intermediate: { graduatesAllCohorts: sumIntermediate },
        advanced: { graduatesAllCohorts: sumAdvanced }
      },
      graduatesAllCohorts: sumGrads,
      natTotalCost: sumCost,
      totalEconomicCostAllCohorts: sumCost,
      epiBenefitAllCohorts: sumBenefit,
      totalEpiBenefitsAllCohorts: sumBenefit,
      netBenefitAllCohorts: sumNet,
      natBcr: (sumCost > 0 ? (sumBenefit / sumCost) : 0),
      _aggregateIndices: indices,
      _aggregateIaGraduates: ia
    };
  }

  const aggregatedShortlistScenario = aggregateShortlistedScenario(shortlist);
  // If the user has added scenarios to the list, the comparison table must
  // present the aggregated package in the Scenario column, regardless of what
  // other tabs are currently rendering.
  const forcedScenarioForComparison = aggregatedShortlistScenario || null;

  const scenarioFromSelector = (() => {
    if (!sel) return null;

    // If the user has added scenarios to the list, aggregate them into a single Scenario column.
    if (aggregatedShortlistScenario) return aggregatedShortlistScenario;

    const v = (appState.baselineCompareActive != null) ? String(appState.baselineCompareActive) : sel.value;
    if (v === "__current__") {
      return appState.currentScenario || null;
    }
    const idx = Number(v);
    if (!isNaN(idx) && saved[idx]) return saved[idx];
    return null;
  })();

  const active = forcedScenarioForComparison || activeScenario || scenarioFromSelector || appState.currentScenario || null;
  if (!active) {
    // Clear the table if nothing is available.
    [
      "baseline-total-grads","baseline-intermediate-grads","baseline-advanced-grads","baseline-frontline-grads","baseline-ia-grads",
      "baseline-remaining-to-target","baseline-total-cost","baseline-total-benefit","baseline-total-net","baseline-total-bcr",
      "scenario-total-grads","scenario-intermediate-grads","scenario-advanced-grads","scenario-frontline-grads","scenario-ia-grads",
      "scenario-remaining-to-target","scenario-total-cost","scenario-total-benefit","scenario-total-net","scenario-total-bcr",
      "delta-total-grads","delta-intermediate-grads","delta-advanced-grads","delta-frontline-grads","delta-ia-grads",
      "delta-remaining-to-target","delta-total-cost","delta-total-benefit","delta-total-net","delta-incremental-bcr"
    ].forEach((id) => setTextIfExists(id, "0"));
    return;
  }

  const baselineTotalGrads = Number(baseline.graduatesAllCohorts || 0);
  const scenarioTotalGrads = Number(active.graduatesAllCohorts || 0);

  const baselineIntermediateGrads = getIntermediateGraduates(baseline) || 0;
  const baselineAdvancedGrads = getAdvancedGraduates(baseline) || 0;
  const baselineFrontlineGrads = getFrontlineGraduates(baseline) || 0;
  const baselineIaGrads = getIntermediateAdvancedGraduates(baseline) || 0;

  const scenarioIntermediateGrads = getIntermediateGraduates(active) || 0;
  const scenarioAdvancedGrads = getAdvancedGraduates(active) || 0;
  const scenarioFrontlineGrads = getFrontlineGraduates(active) || 0;
  const scenarioIaGrads = getIntermediateAdvancedGraduates(active) || 0;

  setTextIfExists("baseline-total-grads", formatNumber(baselineTotalGrads, 0));
  setTextIfExists("baseline-intermediate-grads", formatNumber(baselineIntermediateGrads, 0));
  setTextIfExists("baseline-advanced-grads", formatNumber(baselineAdvancedGrads, 0));
  setTextIfExists("baseline-frontline-grads", formatNumber(baselineFrontlineGrads, 0));
  setTextIfExists("baseline-ia-grads", formatNumber(baselineIaGrads, 0));
  setTextIfExists("baseline-total-cost", formatCurrencyDisplay(baseline.natTotalCost || 0, 0));
  setTextIfExists("baseline-total-benefit", formatCurrencyDisplay(baseline.epiBenefitAllCohorts || 0, 0));
  setTextIfExists("baseline-total-net", formatCurrencyDisplay(baseline.netBenefitAllCohorts || 0, 0));
  setTextIfExists("baseline-total-bcr", baseline.natBcr != null ? formatNumber(baseline.natBcr, 2) : formatNumber(0, 2));

  setTextIfExists("scenario-total-grads", formatNumber(scenarioTotalGrads, 0));
  setTextIfExists("scenario-intermediate-grads", formatNumber(scenarioIntermediateGrads, 0));
  setTextIfExists("scenario-advanced-grads", formatNumber(scenarioAdvancedGrads, 0));
  setTextIfExists("scenario-frontline-grads", formatNumber(scenarioFrontlineGrads, 0));
  setTextIfExists("scenario-ia-grads", formatNumber(scenarioIaGrads, 0));

  // Remaining to reach target (Intermediate + Advanced).
  // Baseline Remaining must use the stock gap defined in the Planner: Target stock − Current stock.
  // Scenario Remaining must reflect: Target stock − (Current stock + scenario Intermediate+Advanced graduates).
  const targetStockEl2 = document.getElementById("planner-target-stock");
  const currentStockEl2 = document.getElementById("planner-current-stock");
  const targetStock = targetStockEl2 ? Number(targetStockEl2.value) : 0;
  const currentStock = currentStockEl2 ? Number(currentStockEl2.value) : 0;

  const baselineRemaining = Math.max(0, targetStock - currentStock);
  // Per user requirement: treat total graduates (Intermediate + Advanced + Frontline) as contributing to closing the target gap.
  const scenarioRemaining = Math.max(0, targetStock - (currentStock + scenarioTotalGrads));
  const deltaRemaining = scenarioRemaining - baselineRemaining;

  setTextIfExists("baseline-remaining-to-target", formatNumber(baselineRemaining, 0));
  setTextIfExists("scenario-remaining-to-target", formatNumber(scenarioRemaining, 0));
  setTextIfExists("delta-remaining-to-target", signedNumber(deltaRemaining, 0));

  const sCost = (active.natTotalCost || active.totalEconomicCostAllCohorts || 0);
  const sBen = (active.epiBenefitAllCohorts || active.totalEpiBenefitsAllCohorts || 0);

  setTextIfExists("scenario-total-cost", formatCurrencyDisplay(sCost, 0));
  setTextIfExists("scenario-total-benefit", formatCurrencyDisplay(sBen, 0));
  setTextIfExists("scenario-total-net", formatCurrencyDisplay(active.netBenefitAllCohorts || (sBen - sCost) || 0, 0));
  setTextIfExists("scenario-total-bcr", active.natBcr != null ? formatNumber(active.natBcr, 2) : formatNumber((sCost > 0 ? (sBen / sCost) : 0), 2));

  setTextIfExists("delta-total-grads", signedNumber(scenarioTotalGrads - baselineTotalGrads, 0));
  setTextIfExists("delta-intermediate-grads", signedNumber(scenarioIntermediateGrads - baselineIntermediateGrads, 0));
  setTextIfExists("delta-advanced-grads", signedNumber(scenarioAdvancedGrads - baselineAdvancedGrads, 0));
  setTextIfExists("delta-frontline-grads", signedNumber(scenarioFrontlineGrads - baselineFrontlineGrads, 0));
  setTextIfExists("delta-ia-grads", signedNumber(scenarioIaGrads - baselineIaGrads, 0));

  setTextIfExists("delta-total-cost", signedCurrencyDisplay(sCost - (baseline.natTotalCost || 0), 0));
  setTextIfExists("delta-total-benefit", signedCurrencyDisplay(sBen - (baseline.epiBenefitAllCohorts || 0), 0));
  setTextIfExists("delta-total-net", signedCurrencyDisplay((active.netBenefitAllCohorts || (sBen - sCost) || 0) - (baseline.netBenefitAllCohorts || 0), 0));

  // Incremental BCR (scenario minus baseline), computed on deltas where meaningful.
  const dCost = sCost - (baseline.natTotalCost || 0);
  const dBen = sBen - (baseline.epiBenefitAllCohorts || 0);
  const incBcr = (dCost > 0) ? (dBen / dCost) : 0;
  setTextIfExists("delta-incremental-bcr", formatNumber(incBcr, 2));
}

function initBaselineEditor() {
  const baselineCfg = loadBaselineConfig();
  appState.baselineConfig = baselineCfg;

  function setVal(id, v) {
    const el = document.getElementById(id);
    if (el) el.value = v;
  }

  // available sites
  setVal('baseline-available-training-sites', Math.max(1, Number((baselineCfg.availableTrainingSites ?? baselineCfg.availableSites) || 1)));

  for (const k of ['frontline','intermediate','advanced']) {
    const t = baselineCfg.tiers[k];
    setVal(`baseline-${k}-cohorts`, t.cohorts);
    setVal(`baseline-${k}-trainees`, t.traineesPerCohort);
    setVal(`baseline-${k}-cost`, t.costPerTraineePerMonth);
    setVal(`baseline-${k}-completion`, t.completionPct);
    setVal(`baseline-${k}-career`, t.career || "certificate");
    setVal(`baseline-${k}-mentorship`, t.mentorship || "medium");
    setVal(`baseline-${k}-delivery`, t.delivery || "blended");
  }

  function readAndPersist() {
    const cfg = loadBaselineConfig();
    const sitesEl = document.getElementById('baseline-available-training-sites');
    let sites = 1;
    if (sitesEl) {
      const v = Number(sitesEl.value);
      sites = isFinite(v) && v >= 1 ? Math.round(v) : 1;
      sitesEl.value = sites;
    }
    cfg.availableTrainingSites = sites;
    cfg.availableSites = sites;

    for (const k of ['frontline','intermediate','advanced']) {
      const cohorts = Number(document.getElementById(`baseline-${k}-cohorts`)?.value);
      const trainees = Number(document.getElementById(`baseline-${k}-trainees`)?.value);
      const cost = Number(document.getElementById(`baseline-${k}-cost`)?.value);
      const comp = Number(document.getElementById(`baseline-${k}-completion`)?.value);
      const careerEl = document.getElementById(`baseline-${k}-career`);
      const mentorshipEl = document.getElementById(`baseline-${k}-mentorship`);
      const deliveryEl = document.getElementById(`baseline-${k}-delivery`);
      const existingTier = cfg.tiers && cfg.tiers[k] ? cfg.tiers[k] : {};

      cfg.tiers[k] = {
        cohorts: Math.max(0, Math.round(isFinite(cohorts) ? cohorts : 0)),
        traineesPerCohort: Math.max(1, Math.round(isFinite(trainees) ? trainees : 1)),
        costPerTraineePerMonth: Math.max(0, isFinite(cost) ? cost : 0),
        completionPct: Math.max(0, Math.min(100, isFinite(comp) ? comp : 0)),
        career: careerEl && careerEl.value ? careerEl.value : (existingTier.career || "certificate"),
        mentorship: mentorshipEl && mentorshipEl.value ? mentorshipEl.value : (existingTier.mentorship || "medium"),
        delivery: deliveryEl && deliveryEl.value ? deliveryEl.value : (existingTier.delivery || "blended")
      };
    }

    persistBaselineConfig(cfg);
    appState.baselineConfig = cfg;

    // refresh everything
    scheduleAutoRecomputeScenario('baseline:change');
    renderScenariosBaselineCompare();
    if (typeof window.renderCopilotPromptBundle === 'function') window.renderCopilotPromptBundle();
  }


const ids = [
    'baseline-available-training-sites',
    'baseline-frontline-cohorts','baseline-frontline-trainees','baseline-frontline-cost','baseline-frontline-completion',
    'baseline-frontline-career','baseline-frontline-mentorship','baseline-frontline-delivery',
    'baseline-intermediate-cohorts','baseline-intermediate-trainees','baseline-intermediate-cost','baseline-intermediate-completion',
    'baseline-intermediate-career','baseline-intermediate-mentorship','baseline-intermediate-delivery',
    'baseline-advanced-cohorts','baseline-advanced-trainees','baseline-advanced-cost','baseline-advanced-completion',
    'baseline-advanced-career','baseline-advanced-mentorship','baseline-advanced-delivery'
  ];

for (const id of ids) {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', readAndPersist);
    if (el) el.addEventListener('change', readAndPersist);
  }

  // buttons
  const copyBtn = document.getElementById('baseline-copy-current');
  if (copyBtn) {
    copyBtn.addEventListener('click', () => {
      const cfg = loadBaselineConfig();
      const current = getConfigFromForm();
      const sitesEl = document.getElementById('available-training-sites');
      let sites = 1;
      if (sitesEl) {
        const v = Number(sitesEl.value);
        sites = isFinite(v) && v >= 1 ? Math.round(v) : 1;
      }
      cfg.availableTrainingSites = sites;

      const tierKey = current.tier;
      if (cfg.tiers[tierKey]) {
        cfg.tiers[tierKey] = {
          cohorts: Math.max(0, Math.round(Number(current.cohorts||0))),
          traineesPerCohort: Math.max(1, Math.round(Number(current.traineesPerCohort||1))),
          costPerTraineePerMonth: Math.max(0, Number(current.costPerTraineePerMonth||0)),
          completionPct: Math.max(0, Math.min(100, Math.round((appState.epiSettings?.general?.completionRateDefault ?? 0.8)*100)))
        };
      }
      persistBaselineConfig(cfg);
      appState.baselineConfig = cfg;
      initBaselineEditor();
      scheduleAutoRecomputeScenario('baseline:copy');
      renderScenariosBaselineCompare();
      if (typeof window.renderCopilotPromptBundle === 'function') window.renderCopilotPromptBundle();
      if (typeof showToast === 'function') showToast("Current configuration copied to baseline.", "success");
    });
  }

  const resetBtn = document.getElementById('baseline-reset');
  if (resetBtn) {
    resetBtn.addEventListener('click', () => {
      const cfg = buildDefaultBaselineConfig();
      persistBaselineConfig(cfg);
      appState.baselineConfig = cfg;
      initBaselineEditor();
      scheduleAutoRecomputeScenario('baseline:reset');
      renderScenariosBaselineCompare();
      if (typeof window.renderCopilotPromptBundle === 'function') window.renderCopilotPromptBundle();
      if (typeof showToast === 'function') showToast("Baseline reset to default.", "success");
    });
  }

  const saveBtn = document.getElementById('baseline-save');
  if (saveBtn) {
    saveBtn.addEventListener('click', () => {
      readAndPersist();
      if (typeof showToast === 'function') showToast("Baseline saved.", "success");
    });
  }

  const applyBtn = document.getElementById('planner-apply-baseline-to-config');
  if (applyBtn) {
    applyBtn.addEventListener('click', () => {
      const cfg = loadBaselineConfig();
      const currentTier = document.getElementById('program-tier')?.value || 'intermediate';
      const t = cfg.tiers[currentTier] || cfg.tiers.intermediate;
      // copy a few comparable controls into current configuration
      const sitesEl = document.getElementById('available-training-sites');
      if (sitesEl) sitesEl.value = Math.max(1, Number(cfg.availableTrainingSites || 1));

      const cohortsEl = document.getElementById('cohorts');
      if (cohortsEl) cohortsEl.value = Math.max(0, Number(t.cohorts || 0));

      const traineesEl = document.getElementById('trainees');
      if (traineesEl) traineesEl.value = Math.max(1, Number(t.traineesPerCohort || 1));

      const costEl = document.getElementById('cost-slider');
      if (costEl) costEl.value = Math.max(0, Number(t.costPerTraineePerMonth || 0));

      const careerEl = document.getElementById('career-track');
      if (careerEl && t.career) careerEl.value = t.career;

      const mentorshipEl = document.getElementById('mentorship');
      if (mentorshipEl && t.mentorship) mentorshipEl.value = t.mentorship;

      const deliveryEl = document.getElementById('delivery-mode');
      if (deliveryEl && t.delivery) deliveryEl.value = t.delivery;

      enforceCohortCapacityConstraints({ tierKey: currentTier, clampIfNeeded: true, source: 'baselineToConfig' });
      scheduleAutoRecomputeScenario('baseline:apply');
      if (typeof showToast === 'function') showToast("Baseline applied to current configuration.", "success");
    });
  }
}

function initPlannerTab() {
  // baseline editor is in planner
  initBaselineEditor();

  const targetStockEl = document.getElementById('planner-target-stock');
  const currentStockEl = document.getElementById('planner-current-stock');
  const targetYearEl = document.getElementById('planner-target-year');
  const baseYearEl = document.getElementById('planner-base-year');

  function getPlannerYears() {
    const targetYear = Number(targetYearEl?.value || 2030);
    const baseYear = Number(baseYearEl?.value || new Date().getFullYear());
    return Math.max(1, Math.round(targetYear - baseYear));
  }

  function recomputePlanner() {
    const years = getPlannerYears();

    // Baseline is defined in the Planner tab baseline editor (BAU counterfactual).
    const baseline = computeBaselineAggregate(years);

    // Scenario is driven by the Configuration tab: baseline tiers + current configured tier override.
    const scenarioAgg = computeScenarioAggregateFromCurrentConfig(years);

    const targetStock = Number(targetStockEl?.value || 7000);
    const currentStock = Number(currentStockEl?.value || 3300);

    const baselineIA = getIntermediateAdvancedGraduates(baseline);
    const scenarioIA = getIntermediateAdvancedGraduates(scenarioAgg);
    const scenarioTotalAllTiers = Number(scenarioAgg.graduatesAllCohorts || scenarioAgg.natTotalGrads || 0);

    // Baseline gap is defined as the difference between the target stock and current stock.
    // This is the reference "Gap to target" shown in the Planner and used as the baseline
    // "Remaining to reach target (Intermediate + Advanced)" in the National simulation tab.
    const baselineGap = Math.max(0, targetStock - currentStock);

    // Per user requirement: treat total graduates (Intermediate + Advanced + Frontline) as contributing to closing the target gap.
    const scenarioRemaining = Math.max(0, targetStock - (currentStock + scenarioTotalAllTiers));
    const annualNeedBaseline = baselineGap / years;

    setTextIfExists('planner-gap', formatNumber(baselineGap, 0));
    setTextIfExists('planner-annual-need', formatNumber(annualNeedBaseline, 1));
    setTextIfExists('planner-baseline-ia', formatNumber(baselineIA, 0));
    setTextIfExists('planner-scenario-ia', formatNumber(scenarioIA, 0));
    setTextIfExists('planner-remaining-gap', formatNumber(scenarioRemaining, 0));

    // Keep baseline scenario up-to-date for compares and prompts
    appState.baselineScenario = baseline;

    // Keep baseline tier endorsements visible in the baseline editor
    try { updateBaselineTierEndorsementsDisplay(baseline); } catch (e) { /* non-fatal */ }

    // Keep National simulation baseline comparison selector up-to-date
    try { renderNationalBaselineScenarioDelta(); } catch (e) { /* non-fatal */ }
}try { appState.recomputePlanner = recomputePlanner; } catch(e) {}

  // Add missing elements defensively
  if (!document.getElementById('planner-baseline-ia')) {
    // optional element may not exist; ignore
  }

  const inputs = [targetStockEl, currentStockEl, targetYearEl, baseYearEl].filter(Boolean);
  inputs.forEach((el) => {
    el.addEventListener('input', recomputePlanner);
    el.addEventListener('change', recomputePlanner);
  });

  const genBtn = document.getElementById('planner-generate-pathways');
  if (genBtn) {
    genBtn.addEventListener('click', () => {
      const years = getPlannerYears();
      const targetStock = Number(targetStockEl?.value || 7000);
      const currentStock = Number(currentStockEl?.value || 3300);
      const baseline = computeBaselineAggregate(years);

      const baselineIA = getIntermediateAdvancedGraduates(baseline);
      // Gap to target is defined as the stock gap (Target - Current), not a projected stock gap.
      const gap = Math.max(0, targetStock - currentStock);

      // Determine per-cohort effective graduates for I and A using current non-scale attributes
      const baseScenarioCfg = getConfigFromForm();

      function perCohortGraduates(tierKey, trainees, completionPct) {
        const completionRateOverride = clamp(Number(completionPct) / 100, 0, 1);
        const cfg = {
          tier: tierKey,
          career: baseScenarioCfg.career,
          mentorship: baseScenarioCfg.mentorship,
          delivery: baseScenarioCfg.delivery,
          response: baseScenarioCfg.response,
          costPerTraineePerMonth: 0,
          traineesPerCohort: trainees,
          cohorts: 1,
          planningHorizonYears: years,
          opportunityCostIncluded: !!baseScenarioCfg.opportunityCostIncluded,
          mentorSupportCostPerCohortBase: Number(baseScenarioCfg.mentorSupportCostPerCohortBase || 0),
          availableMentorsNational: Number(baseScenarioCfg.availableMentorsNational || 0),
          availableTrainingSites: Math.max(1, Number(loadBaselineConfig().availableTrainingSites || 1)),
          maxCohortsPerSitePerYear: Number(baseScenarioCfg.maxCohortsPerSitePerYear || 0),
          crossSectorBenefitMultiplier: Number(baseScenarioCfg.crossSectorBenefitMultiplier || 1.0),
          completionRateOverride
        };
        const s = computeScenario(cfg);
        return Math.max(0, Number(s.graduatesPerCohort || 0));
      }

      const bcfg = loadBaselineConfig();
      const iT = bcfg.tiers.intermediate;
      const aT = bcfg.tiers.advanced;
      const fT = bcfg.tiers.frontline;

      const iPer = perCohortGraduates('intermediate', iT.traineesPerCohort, iT.completionPct);
      const aPer = perCohortGraduates('advanced', aT.traineesPerCohort, aT.completionPct);

      const horizonMonths = years * 12;
      const capI = Math.max(1, Math.floor(horizonMonths / (TIER_MONTHS.intermediate || 12)));
      const capA = Math.max(1, Math.floor(horizonMonths / (TIER_MONTHS.advanced || 24)));

      function buildPathway(name, shareAdvanced, shareIntermediate) {
        const needA = gap * shareAdvanced;
        const needI = gap * shareIntermediate;
        const cohortsA = aPer > 0 ? Math.ceil(needA / aPer) : 0;
        const cohortsI = iPer > 0 ? Math.ceil(needI / iPer) : 0;

        // compute minimum sites needed to accommodate max(cI/capI, cA/capA)
        const sitesNeed = Math.max(
          Math.ceil((cohortsI || 0) / capI),
          Math.ceil((cohortsA || 0) / capA),
          1
        );

        const sites = Math.max(1, Number(bcfg.availableTrainingSites || 1), sitesNeed);

        const portfolioTiers = {
          frontline: {
            tier: 'frontline',
            career: baseScenarioCfg.career,
            mentorship: baseScenarioCfg.mentorship,
            delivery: baseScenarioCfg.delivery,
            response: baseScenarioCfg.response,
            costPerTraineePerMonth: Number(fT.costPerTraineePerMonth || baseScenarioCfg.costPerTraineePerMonth || 0),
            traineesPerCohort: Number(fT.traineesPerCohort || 1),
            cohorts: Math.max(0, Number(fT.cohorts || 0)),
            planningHorizonYears: years,
            opportunityCostIncluded: !!baseScenarioCfg.opportunityCostIncluded,
            mentorSupportCostPerCohortBase: Number(baseScenarioCfg.mentorSupportCostPerCohortBase || 0),
            availableMentorsNational: Number(baseScenarioCfg.availableMentorsNational || 0),
            availableTrainingSites: sites,
            maxCohortsPerSitePerYear: Number(baseScenarioCfg.maxCohortsPerSitePerYear || 0),
            crossSectorBenefitMultiplier: Number(baseScenarioCfg.crossSectorBenefitMultiplier || 1.0),
            completionRateOverride: clamp(Number(fT.completionPct)/100,0,1)
          },
          intermediate: {
            tier: 'intermediate',
            career: baseScenarioCfg.career,
            mentorship: baseScenarioCfg.mentorship,
            delivery: baseScenarioCfg.delivery,
            response: baseScenarioCfg.response,
            costPerTraineePerMonth: Number(iT.costPerTraineePerMonth || baseScenarioCfg.costPerTraineePerMonth || 0),
            traineesPerCohort: Number(iT.traineesPerCohort || 1),
            cohorts: Math.max(0, cohortsI),
            planningHorizonYears: years,
            opportunityCostIncluded: !!baseScenarioCfg.opportunityCostIncluded,
            mentorSupportCostPerCohortBase: Number(baseScenarioCfg.mentorSupportCostPerCohortBase || 0),
            availableMentorsNational: Number(baseScenarioCfg.availableMentorsNational || 0),
            availableTrainingSites: sites,
            maxCohortsPerSitePerYear: Number(baseScenarioCfg.maxCohortsPerSitePerYear || 0),
            crossSectorBenefitMultiplier: Number(baseScenarioCfg.crossSectorBenefitMultiplier || 1.0),
            completionRateOverride: clamp(Number(iT.completionPct)/100,0,1)
          },
          advanced: {
            tier: 'advanced',
            career: baseScenarioCfg.career,
            mentorship: baseScenarioCfg.mentorship,
            delivery: baseScenarioCfg.delivery,
            response: baseScenarioCfg.response,
            costPerTraineePerMonth: Number(aT.costPerTraineePerMonth || baseScenarioCfg.costPerTraineePerMonth || 0),
            traineesPerCohort: Number(aT.traineesPerCohort || 1),
            cohorts: Math.max(0, cohortsA),
            planningHorizonYears: years,
            opportunityCostIncluded: !!baseScenarioCfg.opportunityCostIncluded,
            mentorSupportCostPerCohortBase: Number(baseScenarioCfg.mentorSupportCostPerCohortBase || 0),
            availableMentorsNational: Number(baseScenarioCfg.availableMentorsNational || 0),
            availableTrainingSites: sites,
            maxCohortsPerSitePerYear: Number(baseScenarioCfg.maxCohortsPerSitePerYear || 0),
            crossSectorBenefitMultiplier: Number(baseScenarioCfg.crossSectorBenefitMultiplier || 1.0),
            completionRateOverride: clamp(Number(aT.completionPct)/100,0,1)
          }
        };

        const cfg = {
          isPortfolio: true,
          tier: 'portfolio',
          planningHorizonYears: years,
          availableTrainingSites: sites,
          preferenceModel: baseScenarioCfg.preferenceModel,
          portfolioTiers,
          name,
          notes: 'Generated by Planner: deterministic pathway'
        };

        const scenario = computePortfolioScenario(cfg);
        scenario._sid = randomId('sc');
        scenario.config = cfg;
        scenario.createdAt = Date.now();
        return scenario;
      }

      const pathwayA = buildPathway('Pathway A (Advanced-focused scale-up)', 0.7, 0.3);
      const pathwayB = buildPathway('Pathway B (Intermediate-focused scale-up)', 0.2, 0.8);
      const pathwayC = buildPathway('Pathway C (Balanced scale-up)', 0.4, 0.6);

      // Save and auto-pin
      function upsertScenario(sc) {
        const existing = appState.savedScenarios.find((x) => x.config?.name === sc.config?.name);
        if (existing) {
          // replace config and recalculated fields
          const idx = appState.savedScenarios.indexOf(existing);
          appState.savedScenarios[idx] = sc;
          sc.pinned = true;
          sc.pinnedAt = existing.pinnedAt || Date.now();
        } else {
          sc.pinned = true;
          sc.pinnedAt = Date.now();
          sc.shortlisted = false;
          appState.savedScenarios.push(sc);
        }
      }

      upsertScenario(pathwayA);
      upsertScenario(pathwayB);
      upsertScenario(pathwayC);

      enforcePinLimitDeterministic();
      persistSavedScenarios();
      refreshSavedScenariosTable();
      refreshTopScenariosPanel('top5');
      if (typeof showToast === 'function') showToast("Pathways A, B, and C generated and saved.", "success");
      refreshTopScenariosPanel('top5-copilot');
      renderScenariosBaselineCompare();
      if (typeof window.renderCopilotPromptBundle === 'function') window.renderCopilotPromptBundle();
    });
  }

  // initial compute
  recomputePlanner();

  // keep planner aligned when baseline changes via editor
  const baselineWrap = document.getElementById('plannerTab');
  if (baselineWrap) {
    baselineWrap.addEventListener('input', () => {
      recomputePlanner();
    });
  }
}


document.addEventListener("DOMContentLoaded", () => {
  COST_CONFIG = COST_TEMPLATES;

  initTabs();
  initDefinitionTooltips();
  initTooltips();
  initGuidedTour();
  initAdvancedSettings();
      initButtonToasts();
  initCapacityCostLogging();
  initCopilot();
  initCopilotPromptButtonFallback();

  try { if (typeof initPlannerTab === 'function') initPlannerTab(); } catch (e) { /* non-fatal */ }
  // Brief export bullets preview
  renderBriefBulletsPreview();
  const _en = document.getElementById("export-enablers");
  const _rk = document.getElementById("export-risks");
  if (_en) _en.addEventListener("input", renderBriefBulletsPreview);
  if (_rk) _rk.addEventListener("input", renderBriefBulletsPreview);

  // Initial Top 5 renders
  refreshTopScenariosPanel("top5");
  refreshTopScenariosPanel("top5-copilot");

  enforceResponseTimeFixedTo7Days();
  initOutbreakSensitivityDropdowns();

  initEventHandlers();
  enforceCohortCapacityConstraints({ tierKey: (document.getElementById("program-tier") ? document.getElementById("program-tier").value : "frontline"), clampIfNeeded: true, source: "init" });

  // STEPS UPGRADE: status quo copy button
  const copyBtn = document.getElementById("copy-status-quo");
  if (copyBtn) copyBtn.addEventListener("click", () => copyStatusQuoToClipboard());

  updateCostSliderLabel();
  updateCurrencyToggle();
});

function initButtonToasts() {
      document.body.addEventListener("click", (ev) => {
        const target = ev.target;
        if (!target) return;
        const el = target.closest("[data-toast-message]");
        if (!el) return;
        const msg = el.getAttribute("data-toast-message");
        if (!msg) return;
        const t = el.getAttribute("data-toast-type") || "info";
        showToast(msg, t);
      });
    }





function exportSensitivityBenefitsTableToWord(){
  if (typeof exportSensitivityBenefitsToPdf === "function") return exportSensitivityBenefitsToPdf();
  if (typeof exportSensitivityBenefitsPDF === "function") return exportSensitivityBenefitsPDF();
  if (typeof exportSensitivityTableToWord === "function") return exportSensitivityTableToWord();
  throw new Error("Sensitivity export function not found.");
}

/* EXPORT_BUTTON_FALLBACKS */
(function attachExportButtonFallbacks(){
  if (window.__stepsExportFallbacksAttached) return;
  window.__stepsExportFallbacksAttached = true;
  document.addEventListener("click", function(ev){
    const t = ev.target && ev.target.closest ? ev.target.closest("button") : null;
    if (!t || !t.id) return;
    const id = t.id;
    if (id === "export-pdf") { ev.preventDefault(); try { exportScenariosToPdf(); } catch (e) { try { showToast("Export failed: " + (e && e.message ? e.message : "unknown error"), "error"); } catch (_) {} } return; }
    if (id === "export-top5-pdf") { ev.preventDefault(); try { exportTop5OnlyPdf(); } catch (e) { try { showToast("Export failed: " + (e && e.message ? e.message : "unknown error"), "error"); } catch (_) {} } return; }
    if (id === "export-excel") { ev.preventDefault(); try { exportScenariosToExcel(); } catch (e) { try { showToast("Export failed: " + (e && e.message ? e.message : "unknown error"), "error"); } catch (_) {} } return; }
    if (id === "export-top5-excel") { ev.preventDefault(); try { exportTop5OnlyExcel(); } catch (e) { try { showToast("Export failed: " + (e && e.message ? e.message : "unknown error"), "error"); } catch (_) {} } return; }
    if (id === "downloadSensitivityPDF") { ev.preventDefault(); try { exportSensitivityBenefitsTableToWord(); } catch (e) { try { showToast("Export failed: " + (e && e.message ? e.message : "unknown error"), "error"); } catch (_) {} } return; }
  }, true);
})();
