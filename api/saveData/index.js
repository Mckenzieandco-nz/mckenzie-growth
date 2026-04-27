const { database } = require('../shared/cosmos');

module.exports = async function (context, req) {
  try {
    const data = req.body;
    if (!data || typeof data !== 'object') {
      context.res = { status: 400, body: { error: 'Request body must be a JSON object' } };
      return;
    }

    const collections = ['users', 'materials', 'assignments', 'reviews', 'onetoones', 'cpdLog', 'disclosures', 'docs', 'kpis'];

    for (const name of collections) {
      if (data[name] !== undefined) {
        const container = database.container(name);
        await container.items.upsert({ id: 'all', data: data[name] });
      }
    }

    // Store skillsCatalog and any other settings in the settings container
    const settingsData = {};
    if (data.skillsCatalog !== undefined) settingsData.skillsCatalog = data.skillsCatalog;
    if (data.resources !== undefined) settingsData.resources = data.resources;
    if (data.technicalExperts !== undefined) settingsData.technicalExperts = data.technicalExperts;
    if (data.disclosureContact !== undefined) settingsData.disclosureContact = data.disclosureContact;
    if (data.reportRecipients !== undefined) settingsData.reportRecipients = data.reportRecipients;
    if (data.kpiAutoMonthly !== undefined) settingsData.kpiAutoMonthly = data.kpiAutoMonthly;
    if (data.kpiAutoQuarterly !== undefined) settingsData.kpiAutoQuarterly = data.kpiAutoQuarterly;
    if (data.kpiAutoAnnual !== undefined) settingsData.kpiAutoAnnual = data.kpiAutoAnnual;
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
