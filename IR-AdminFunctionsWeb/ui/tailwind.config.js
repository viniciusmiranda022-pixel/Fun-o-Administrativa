/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      colors: {
        sidebar: '#1F1F1F',
        'sidebar-sub': '#3A3A3A',
        topbar: '#252525',
        panel: '#F5F5F5',
        quest: { blue: '#0096D6', dark: '#0078A8', hover: '#006090' },
        status: { success: '#28A745', error: '#DC3545', warning: '#FFC107', info: '#17A2B8' },
        border: { light: '#DEE2E6', medium: '#C8CDD3' }
      },
      fontFamily: { sans: ['"Segoe UI"', 'system-ui', 'sans-serif'] }
    }
  },
  plugins: []
};
