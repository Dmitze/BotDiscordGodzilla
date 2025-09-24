/** @type {import('dependency-cruiser').IConfiguration} */
module.exports = {
  forbidden: [
    // no circulars
    { name: 'no-circular', severity: 'error', from: {}, to: { circular: true } },
    // layering rules (adapters can depend on core & services; core on services; services on infra; infra on utils/types)
    {
      name: 'layering-adapters',
      comment: 'Adapters (api|chat|commands|ui) must not import infra directly',
      severity: 'error',
      from: { path: '^src/(api|chat|commands|ui)/' },
      to: { path: '^src/(search|workspace)/' }
    },
    {
      name: 'layering-core',
      comment: 'Core should not depend on adapters',
      severity: 'error',
      from: { path: '^src/core/' },
      to: { path: '^src/(api|chat|commands|ui)/' }
    }
  ],
  options: {
    doNotFollow: { path: 'node_modules|^\.\.|^..' },
    tsPreCompilationDeps: true,
    combinedDependencies: true,
    reporterOptions: { dot: { collapsePattern: 'node_modules/.*' } }
  }
};


