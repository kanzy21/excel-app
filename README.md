# React Excel Application

This is a simple excel application built with React and Vite, featuring functionalities like editable cells, dynamic rows and columns, formula calculations (`SUM`, `AVERAGE`), cell formatting (bold and background color), and saving/loading data in JSON format.

## Features
- Display a 10x10 grid of editable cells.
- Dynamically add rows and columns.
- Perform `SUM` and `AVERAGE` calculations using formulas.
- Format cells with bold text and background color.
- Copy and paste selected ranges of cells. (WIP)
- Save the spreadsheet as a JSON file and load it back.

## Requirements
- Node.js (version 14 or higher)
- npm or yarn

## Getting Started

### 1. Clone the Repository
```bash
git clone https://github.com/your-username/react-spreadsheet-app.git
cd excel-app
```
## Cloning and Running the Application in local

Clone the project into local

Install all the npm packages. Go into the project folder and type the following command to install all npm packages

```bash
npm install
```

In order to run the application Type the following command

```bash
npm run dev
```

The Application Runs on **http://127.0.0.1:5174**

This will directly open the application and you can use the features mentioned above except `copy` and `paste`.

To use feature of sum and average calculation you have to write `=AVERAGE(A1:A3)` or `=SUM(A1:A3)` where A1 is 1st cell of 1st row and 1st column and A3 is 1st cell of 1st column and 3rd row.

