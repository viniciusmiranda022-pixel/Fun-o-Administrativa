/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      colors: {
        sidebar: '#1F1F1F',
        panel: '#F8F9FA',
        accent: {
          DEFAULT: '#0066CC',
          hover: '#0052A3'
        },
        status: {
          success: '#28A745',
          error: '#DC3545',
          warning: '#FFC107'
        },
        border: {
          light: '#DEE2E6'
        }
      },
      fontFamily: {
        sans: ['"Segoe UI"', 'system-ui', 'sans-serif']
      }
    }
  },
  plugins: []
};
