const { database } = require('../shared/cosmos');

module.exports = async function (context, req) {
  try {
    const data = req.body;
    if (!data || typeof data !== 'object') {
      context.res = { status: 400, body: { error: 'Request body must be a JSON object' } };
      return;
    }

    // Core data collections — each wrapped individually so one failure doesn't block the rest
    const collections = ['users', 'materials', 'assignments', 'reviews', 'onetoones', 'cpdLog', 'disclosures', 'docs'];

    for (const name of collections) {
      if (data[name] !== undefined) {
        try {
          const container = database.container(name);
          await container.items.upsert({ id: 'all', data: data[name] });
        } catch (e) {
          context.log.warn(`saveData: failed to save collection '${name}':`, e.message);
        }
      }
    }

    // Settings document — stores config/flags that don't need their own container
    // kpis, featureFlags, reportRecipients, etc. all live here alongside skillsCatalog
    const settingsData = {};
    if (data.skillsCatalog !== undefined) settingsData.skillsCatalog = data.skillsCatalog;
    if (data.resources !== undefined) settingsData.resources = data.resources;
    if (data.technicalExperts !== undefined) settingsData.technicalExperts = data.technicalExperts;
    if (data.disclosureContact !== undefined) settingsData.disclosureContact = data.disclosureContact;
    if (data.kpis !== undefined) settingsData.kpis = data.kpis;
    if (data.reportRecipients !== undefined) settingsData.reportRecipients = data.reportRecipients;
    if (data.kpiAutoMonthly !== undefined) settingsData.kpiAutoMonthly = data.kpiAutoMonthly;
    if (data.kpiAutoQuarterly !== undefined) settingsData.kpiAutoQuarterly = data.kpiAutoQuarterly;
    if (data.kpiAutoAnnual !== undefined) settingsData.kpiAutoAnnual = data.kpiAutoAnnual;
    if (data.featureFlags !== undefined) settingsData.featureFlags = data.featureFlags;
    if (Object.keys(settingsData).length > 0) {
      const settingsContainer = database.container('settings');
      await settingsContainer.items.upsert({ id: 'all', data: settingsData });
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true }
    };
  } catch (err) {
    context.log.error('saveData error:', err.message);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message }
    };
  }
};
