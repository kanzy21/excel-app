import React, { useState } from 'react';
import './excel.css';

const initialRows = 10;
const initialCols = 10;
const Excel = () => {

	const [data, setData] = useState(Array.from({ length: initialRows }, () => Array(initialCols).fill('')));
	const [rows, setRows] = useState(initialRows);
	const [cols, setCols] = useState(initialCols);
	const [selectedCell, setSelectedCell] = useState({ row: null, col: null });
	const [selectionStart, setSelectionStart] = useState(null);
	const [selectionEnd, setSelectionEnd] = useState(null);
	const [copiedCells, setCopiedCells] = useState([]);
	const [boldCells, setBoldCells] = useState({});
	const [cellColors, setCellColors] = useState({});
	const [formulaCells, setFormulaCells] = useState({});

	// Parse and calculate formulas
	const calculateFormula = (value, rowIndex, colIndex) => {
		if (rowIndex >= 0 && colIndex >= 0) {
			if (value.startsWith('=SUM(')) {
				const range = value.match(/\((.*)\)/)?.length >= 1 ? value.match(/\((.*)\)/)[1] : [];
				if (range[1]) {
					const [start, end] = range.split(':');
					if (start && end) {
						const result = sumRange(start, end);
						if (start && end && typeof sumRange(start, end) == 'number') {
							let temp = formulaCells;
							temp[rowIndex + '-' + colIndex] = result;
							setFormulaCells({ ...temp });
							return result;
						}
					}
				}
				// return sumRange(start, end);
			}
			if (value.startsWith('=AVERAGE(')) {
				const range = value.match(/\((.*)\)/)?.length >= 1 ? value.match(/\((.*)\)/)[1] : [];
				if (range[1]) {
					const [start, end] = range.split(':');
					if (start && end) {
						const result = averageRange(start, end);
						if (start && end && typeof averageRange(start, end) == 'number') {
							let temp = formulaCells;
							temp[rowIndex + '-' + colIndex] = result;
							setFormulaCells({ ...temp });
							return result;
						}
					}
				}
				// return averageRange(start, end);
			}
		}
		return value;
	};

	// Sum of cells
	const sumRange = (start, end) => {
		const [startRow, startCol] = parseCell(start);
		const [endRow, endCol] = parseCell(end);
		let sum = 0;
		for (let row = startRow; row <= endRow; row++) {
			for (let col = startCol; col <= endCol; col++) {
				sum += Number(data[row][col]) || 0;
			}
		}
		return sum;
	};

	// Average of cells
	const averageRange = (start, end) => {
		const [startRow, startCol] = parseCell(start);
		const [endRow, endCol] = parseCell(end);
		let sum = 0,
			count = 0;
		for (let row = startRow; row <= endRow; row++) {
			for (let col = startCol; col <= endCol; col++) {
				sum += Number(data[row][col]) || 0;
				count++;
			}
		}
		return count ? sum / count : 0;
	};

	// parse A1 or B2 etc
	const parseCell = (cell) => {
		const col = cell.charCodeAt(0) - 65; // A -> 0, B -> 1, etc.
		const row = parseInt(cell.slice(1)) - 1;
		return [row, col];
	};

	// Handle cell value change
	const handleChange = (rowIndex, colIndex, value) => {
		const updatedData = [...data];
		let val = value;
		if (formulaCells[rowIndex + '-' + colIndex]) {
			let temp = formulaCells;
			delete temp[rowIndex + '-' + colIndex];
			setFormulaCells({ ...temp });
		}
		if (value?.startsWith('=')) {
			val = calculateFormula(value, rowIndex, colIndex);
		}

		updatedData[rowIndex][colIndex] = val;
		setData(updatedData);
	};

	// Selection handling for copy-paste (This is not working)
	const handleMouseDown = (rowIndex, colIndex) => {
		setSelectionStart({ row: rowIndex, col: colIndex });
		setSelectionEnd({ row: rowIndex, col: colIndex });
	};

	const handleMouseEnter = (rowIndex, colIndex) => {
		if (selectionStart) {
			setSelectionEnd({ row: rowIndex, col: colIndex });
		}
	};

	const handleMouseUp = () => {
		setSelectionStart(null);
	};

	const isSelected = (rowIndex, colIndex) => {
		if (!selectionStart || !selectionEnd) return false;

		const [minRow, maxRow] = [selectionStart.row, selectionEnd.row].sort((a, b) => a - b);
		const [minCol, maxCol] = [selectionStart.col, selectionEnd.col].sort((a, b) => a - b);

		return rowIndex >= minRow && rowIndex <= maxRow && colIndex >= minCol && colIndex <= maxCol;
	};

	// Copy and Paste handling (This is not working)
	const handleCopy = () => {
		if (!selectionStart || !selectionEnd) return;

		const copiedData = [];
		const [minRow, maxRow] = [selectionStart.row, selectionEnd.row].sort((a, b) => a - b);
		const [minCol, maxCol] = [selectionStart.col, selectionEnd.col].sort((a, b) => a - b);

		for (let row = minRow; row <= maxRow; row++) {
			const rowCopy = [];
			for (let col = minCol; col <= maxCol; col++) {
				rowCopy.push(data[row][col]);
			}
			copiedData.push(rowCopy);
		}

		setCopiedCells(copiedData);
	};
	//(This is not working)
	const handlePaste = () => {
		if (!selectedCell || copiedCells.length === 0) return;

		const newData = [...data];
		const numRows = copiedCells.length;
		const numCols = copiedCells[0].length;

		for (let row = 0; row < numRows; row++) {
			for (let col = 0; col < numCols; col++) {
				const targetRow = selectedCell.row + row;
				const targetCol = selectedCell.col + col;
				if (targetRow < rows && targetCol < cols) {
					newData[targetRow][targetCol] = copiedCells[row][col];
				}
			}
		}

		setData(newData);
	};

	// Add row
	const addRow = () => {
		setData([...data, Array(cols).fill('')]);
		setRows(rows + 1);
	};
	// Add Column
	const addColumn = () => {
		const newData = data.map((row) => [...row, '']);
		setData(newData);
		setCols(cols + 1);
	};

	// Save to JSON
	const saveToJson = () => {
		const json = JSON.stringify({ data, boldCells, cellColors });
		const blob = new Blob([json], { type: 'application/json' });
		const url = URL.createObjectURL(blob);
		const link = document.createElement('a');
		link.href = url;
		link.download = 'spreadsheet.json';
		link.click();
	};

	// Load from JSON
	const loadFromJson = (file) => {
		const reader = new FileReader();
		reader.onload = (e) => {
			const jsonData = JSON.parse(e.target.result);
			setData(jsonData.data);
			setBoldCells(jsonData.boldCells || {});
			setCellColors(jsonData.cellColors || {});
		};
		reader.readAsText(file);
	};

	// Bold handling
	const handleBold = () => {
		const cellKey = `${selectedCell.row}-${selectedCell.col}`;
		setBoldCells((prev) => ({
			...prev,
			[cellKey]: !prev[cellKey],
		}));
	};

	// Color Handler
	const handleColorChange = (color) => {
		const cellKey = `${selectedCell.row}-${selectedCell.col}`;
		setCellColors((prev) => ({
			...prev,
			[cellKey]: color,
		}));
	};
	return (
		<div className="excel-container">
			<div className="controls">
				<button onClick={addRow}>Add Row</button>
				<button onClick={addColumn}>Add Column</button>
				<button onClick={handleBold}>Bold</button>
				<input type="color" onChange={(e) => handleColorChange(e.target.value)} />
				<button onClick={saveToJson}>Save</button>
				<input type="file" accept=".json" onChange={(e) => loadFromJson(e.target.files[0])} />
				<button onClick={handleCopy}>Copy</button>
				<button onClick={handlePaste}>Paste</button>
			</div>

			<table className="excel-table">
				<tbody>
					{data.map((row, rowIndex) => (
						<tr key={rowIndex}>
							{row.map((cell, colIndex) => (
								<td key={colIndex}>
									<input
										type="text"
										value={formulaCells[`${rowIndex}-${colIndex}`] || cell}
										onChange={(e) => handleChange(rowIndex, colIndex, e.target.value)}
										onMouseDown={() => setSelectedCell({ row: rowIndex, col: colIndex })}
										onMouseEnter={() => handleMouseEnter(rowIndex, colIndex)}
										onMouseUp={handleMouseUp}
										style={{
											backgroundColor: isSelected(rowIndex, colIndex)
												? 'lightblue'
												: cellColors[`${rowIndex}-${colIndex}`] || 'white',
											fontWeight: boldCells[`${rowIndex}-${colIndex}`] ? 'bold' : 'normal',
										}}
									/>
								</td>
							))}
						</tr>
					))}
				</tbody>
			</table>
		</div>
	);
};

export default Excel;
