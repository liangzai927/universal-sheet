export default [
  {
    ignores: ['node_modules', 'dist', '.turbo', 'coverage', '.vitepress'],
  },
  {
    files: ['**/*.ts'],
    languageOptions: {
      parserOptions: {
        project: true,
      },
    },
    rules: {
      'no-console': 'warn',
      'no-unused-vars': 'off',
    },
  },
  {
    files: ['**/*.js', '**/*.mjs'],
    rules: {
      'no-console': 'off',
    },
  },
];
