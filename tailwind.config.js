/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        nexus: {
          bg: '#f8fafc',
          card: '#ffffff',
          primary: '#0284c7', // calm blue
          'primary-hover': '#0369a1',
          teal: '#0d9488',
          green: '#16a34a',
          dark: '#0f172a',
          muted: '#64748b',
          border: '#e2e8f0',
        },
        alert: {
          safe: '#22c55e',
          caution: '#eab308',
          warning: '#f97316',
          critical: '#ef4444',
          info: '#3b82f6',
        }
      },
      fontFamily: {
        sans: ['Inter', 'system-ui', 'sans-serif'],
      },
      boxShadow: {
        'nexus-sm': '0 1px 3px rgba(0,0,0,0.05), 0 1px 2px rgba(0,0,0,0.1)',
        'nexus-md': '0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03)',
        'nexus-lg': '0 10px 15px -3px rgba(0, 0, 0, 0.07), 0 4px 6px -2px rgba(0, 0, 0, 0.04)',
        'nexus-glow': '0 0 20px rgba(2, 132, 199, 0.25)',
        'alert-glow': '0 0 20px rgba(239, 68, 68, 0.35)',
      }
    },
  },
  plugins: [],
}
