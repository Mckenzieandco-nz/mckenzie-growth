const { database } = require('../shared/cosmos');

module.exports = async function (context, req) {
  try {
    const collections = ['users', 'materials', 'assignments', 'reviews', 'onetoones', 'cpdLog', 'disclosures', 'docs', 'kpis', 'settings'];
    const result = {};

    for (const name of collections) {
      const container = database.container(name);
      try {
        const { resource } = await container.item('all', 'all').read();
        result[name] = resource ? resource.data : (name === 'settings' ? {} : []);
      } catch (e) {
        // Container might be empty — that's fine
        result[name] = [];
      }
    }

    // Flatten settings into top-level keys
    const settings = result.settings || {};
    delete result.settings;
    Object.assign(result, settings);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: result
    };
  } catch (err) {
    context.log.error('getData error:', err.message);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message }
    };
  }
};
