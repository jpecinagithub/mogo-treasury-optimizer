/** @type {import('tailwindcss').Config} */
export default {
  content: [
    './index.html',
    './src/**/*.{js,ts,jsx,tsx}',
  ],
  theme: {
    extend: {
      fontFamily: { sans: ['Inter','sans-serif'], mono: ['JetBrains Mono','monospace'] },
      colors: {
        mogo: { navy:'#0B1E3A', blue:'#162E5B', light:'#FFF7ED', orange:'#FF6B00', yellow:'#FCDC04', red:'#D90000', accent:'#FF6B00', teal:'#0d9488' },
      },
    },
  },
  plugins: [],
}

