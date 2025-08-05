import type { Config } from 'lint-staged';

const config: Config = {
  '*.{js,ts}': [
    'eslint --fix',
    'prettier --write',
  ],
  '*.{json,md,yml,yaml}': [
    'prettier --write',
  ],
};

export default config; 