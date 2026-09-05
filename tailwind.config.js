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
        }
      }
    },
  },
  plugins: [],
}
