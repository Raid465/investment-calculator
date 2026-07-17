<div align="center">

# Investment Calculator

### Interactive Compound Interest Calculator

</div>

<div align="center">

![React](https://img.shields.io/badge/React-19-61DAFB?style=flat-square&logo=react)
![TypeScript](https://img.shields.io/badge/TypeScript-6-3178C6?style=flat-square&logo=typescript)
![Vite](https://img.shields.io/badge/Vite-8-646CFF?style=flat-square&logo=vite)
![Tailwind](https://img.shields.io/badge/Tailwind-4-06B6D4?style=flat-square&logo=tailwindcss)
![Recharts](https://img.shields.io/badge/Recharts-3-8884d8?style=flat-square)

</div>

---

A comprehensive interactive tool for investment planning and compound interest calculation, with dynamic charts and professional Excel report export.

## Features

### Compound Interest Calculator
- Initial capital and monthly deposit inputs
- Expected annual return rate and investment duration
- Final value, total earnings, and growth percentage display
- Inflation adjustment (real value)

### Interactive Charts
- Stacked Area chart separating invested amount from earnings
- Interactive tooltips with percentage breakdown
- K/M number formatting for easy reading

### Year-by-Year Table
- Cumulative invested amount
- Portfolio value each year
- Cumulative earnings and growth rate
- Real value when inflation is enabled

### Scenario Comparison
- 3 scenarios: Conservative (10%), Moderate (18%), Optimistic (35%)
- Shared inputs (capital, deposit, duration)
- Comparative chart for all scenarios
- Detailed comparison table

### Reverse Goal Calculator
- Set target amount and target year
- Calculate required monthly payment to reach goal
- Detect if goal is already achievable
- Growth projection chart with goal reference line

### Professional Excel Export
- `.xlsx` file with dark mode design
- Formatted numbers with thousand separators and currency symbols
- Color-coded columns with gradients
- Embedded chart in the file
- Multiple sheets for each scenario

### Currency Support
- US Dollar ($)
- Jordanian Dinar (د.أ)

## Tech Stack

| Technology | Usage |
|---|---|
| **React 19** | Interactive UI |
| **TypeScript** | Type-safe code |
| **Vite** | Fast build and dev server |
| **Tailwind CSS** | Dark mode styling |
| **Recharts** | Interactive charts |
| **ExcelJS** | Excel report export |
| **file-saver** | File downloads |

## Getting Started

### Prerequisites
- Node.js 18+
- npm or yarn

### Installation

```bash
# Clone the repository
git clone https://github.com/Raid465/investment-calculator.git
cd investment-calculator

# Install dependencies
npm install

# Start dev server
npm run dev

# Build for production
npm run build
```

## Project Structure

```
src/
├── App.tsx          # Main component with all features
├── main.tsx         # Entry point
└── index.css        # Tailwind base styles
```

## License

This project is open source and available for any use.

---

<div align="center">

**All calculations are approximate and do not constitute financial advice.**

</div>
