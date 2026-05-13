import { defineConfig } from 'vitepress';

export default defineConfig({
  title: 'Universal Sheet',
  description: 'A framework-agnostic spreadsheet SDK',
  themeConfig: {
    nav: [
      { text: 'Guide', link: '/guide/' },
      { text: 'API', link: '/api/' },
    ],
    sidebar: {
      '/guide/': [
        { text: 'Getting Started', link: '/guide/' },
        { text: 'Core Concepts', link: '/guide/concepts' },
      ],
      '/api/': [
        { text: 'Core', link: '/api/core' },
        { text: 'Engine', link: '/api/engine' },
        { text: 'Formula', link: '/api/formula' },
      ],
    },
    socialLinks: [
      { icon: 'github', link: 'https://github.com' },
    ],
  },
});
