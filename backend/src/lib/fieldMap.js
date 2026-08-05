// Mirrors FIELD_MAP / DATE_FIELDS / spItemToRecord / recordToSpFields in the
// frontend's app.js — must stay identical or record shapes will drift between
// the two. Once the frontend is cut over to call this backend (Phase 5), the
// frontend copies can be deleted and this becomes the only copy.
const FIELD_MAP = {
  clients: { client_name: 'Title', company: 'field_1', designation: 'field_2', department: 'field_3', therapy_area: 'field_4', region: 'field_5', email: 'field_6', phone: 'field_7', linkedin_url: 'field_8', status: 'field_9', priority: 'field_10', assigned_bd: 'field_11', notes: 'field_12' },
  opportunities: { opportunity: 'Title', rfp_id: 'field_1', company: 'field_2', client_name: 'field_3', opportunity_status: 'field_4', discussion_date: 'field_5', pitch_date: 'field_6', proposal_submission_date: 'field_7', expected_close_date: 'field_8', bd_owner: 'field_9', supporting_role: 'field_10', estimated_value_usd: 'field_11', probability_pct: 'field_12', probability_weighted_value: 'field_13', notes: 'field_14', stage_history: 'StageHistory', identified_date: 'IdentifiedDate', fx_rate_locked: 'FXRateLocked' },
  engagements: { client_name: 'Title', eng_month: 'field_1', eng_date: 'field_2', designation: 'field_3', company: 'field_4', bd_pm: 'field_5', engagement_type: 'field_6', stakeholder_type: 'field_8', engagement_objective: 'field_9', engagement_outcome: 'field_10', discussion_points: 'field_11', cta_next_step: 'field_12', cta_due_date: 'field_13', cta_owner: 'field_14', follow_up_done: 'field_15', accompanied_by: 'AccompaniedBy' },
  companies: { company: 'Title', onboarding_status: 'field_1', notes: 'field_2', target_revenue: 'TargetRevenue', overall_budget_potential: 'BudgetPotential', overall_client_relationship: 'ClientRelationship', client_perception: 'ClientPerception', team_satisfaction: 'TeamSatisfaction', degree_of_innovation: 'DegreeOfInnovation' }
};

const DATE_FIELDS = {
  clients: [],
  opportunities: ['discussion_date', 'pitch_date', 'proposal_submission_date', 'expected_close_date', 'identified_date'],
  engagements: ['eng_month', 'eng_date', 'cta_due_date'],
  companies: []
};

function spItemToRecord(kind, item) {
  const map = FIELD_MAP[kind], f = item.fields || {}, rec = { id: item.id };
  Object.entries(map).forEach(([k, sp]) => {
    let v = f[sp];
    if (v != null && DATE_FIELDS[kind].includes(k)) v = String(v).split('T')[0];
    rec[k] = v ?? null;
  });
  return rec;
}

function recordToSpFields(kind, payload) {
  const map = FIELD_MAP[kind], fields = {};
  Object.entries(payload).forEach(([k, v]) => {
    const sp = map[k]; if (!sp) return;
    if (v != null && DATE_FIELDS[kind].includes(k)) v = new Date(v).toISOString();
    fields[sp] = v;
  });
  return fields;
}

module.exports = { FIELD_MAP, DATE_FIELDS, spItemToRecord, recordToSpFields };
