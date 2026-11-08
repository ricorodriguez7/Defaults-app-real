import type { Config } from 'tailwindcss';

const config: Config = {
  content: ['./app/**/*.{js,ts,jsx,tsx,mdx}'],
  theme: {
    extend: {
      colors: {
        navy: '#172d50',
        accent: '#4268de',
      },
    },
  },
  plugins: [],
};

export default config;
