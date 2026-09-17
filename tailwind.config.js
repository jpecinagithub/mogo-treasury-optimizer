/** @type {import('tailwindcss').Config} */
export default {
  content: [
    './index.html',
    './src/**/*.{js,ts,jsx,tsx}',
  ],
  theme: {
    extend: {
      fontFamily: { sans: ['Inter','system-ui','sans-serif'], mono: ['JetBrains Mono','monospace'] },
      colors: {
        mogo: {
          navy: '#0B1E3A',
          navyLight: '#13294B',
          slate: '#F8FAFC',
          border: '#E2E8F0',
          gold: '#C6A664',
          goldLight: '#F0E6CC',
          teal: '#0FA3A8',
          // legacy aliases kept for safety
          blue: '#162E5B',
          light: '#FFF7ED',
          orange: '#C6A664',
          yellow: '#FCDC04',
          red: '#D90000',
          accent: '#C6A664',
        },
      },
      boxShadow: {
        card: '0 1px 3px #0f1e3a14, 0 4px 12px #0f1e3a0f',
        elevated: '0 8px 30px #0f1e3a1f',
      },
    },
  },
  plugins: [],
}
