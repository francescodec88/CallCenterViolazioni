module.exports = {
  content: [
    './src/**/*.{html,ts,scss}',
    './src/**/*.component.html',
    './src/**/*.component.ts'
  ],
  theme: {
    extend: {
      colors: {
        primary: '#3B82F6', // blue-500
        secondary: '#10B981', // emerald-500
        accent: '#F59E0B' // amber-500
      },
      backdropBlur: {
        xs: '2px'
      }
    }
  },
  plugins: []
};
