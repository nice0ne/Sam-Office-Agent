/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        office: {
          excel: '#107C41',
          word: '#185ABD',
          ppt: '#C43E1C',
          agent: '#0F6CBD',
        },
        gray: {
          750: '#232d3d',
          850: '#141b26',
          950: '#0b0f17',
        },
        surface: {
          light: '#ffffff',
          dark: '#111827',
          'card-light': '#ffffff',
          'card-dark': '#1e293b',
        },
      },
      boxShadow: {
        '2xs': '0 1px 2px 0 rgba(0, 0, 0, 0.05)',
        'dark-card': '0 1px 3px 0 rgba(0, 0, 0, 0.4), 0 1px 2px -1px rgba(0, 0, 0, 0.4)',
      },
    },
  },
  plugins: [],
}
