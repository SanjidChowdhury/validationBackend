"use strict"
import ExcelJS from 'exceljs';
import Express from 'express';
import cors from 'cors';
import bodyParser from 'body-parser';
import path from 'path';

const excel_files_path = path.join(process.cwd(), 'ExcelFiles/')
const app = Express();
const port = 3000

const dataset_file_path = path.join(excel_files_path + 'dataset.xlsx');
const dataset_workbook = new ExcelJS.Workbook();

const validation_file_path = path.join(excel_files_path + 'validation.xlsx');
const validation_workbook = new ExcelJS.Workbook();

const total = 25; 



let corsOptions = {
	origin: '*'
  }

app.use(cors(corsOptions))

app.use(bodyParser.json())

app.get('/', (req, res) => {
	res.send('hello world')
})

app.get('/data', (req, res) => {
	let id = parseInt(req.query.id);
	let no_of_rows = total;
	id = id * no_of_rows;
	id = id + 2;

	dataset_workbook.xlsx.readFile(dataset_file_path).then(() => {
		const worksheet = dataset_workbook.worksheets[0];
		const rows = worksheet.getRows(parseInt(id), no_of_rows)

		let data = []

		rows.forEach((row) => {
			data.push(row.getCell(1).value)
		})

		res.send(data);
	})
})

app.get('/choices', (req, res) => {
	dataset_workbook.xlsx.readFile(dataset_file_path).then(() => {
		const worksheet = dataset_workbook.worksheets[0];
		const firstRow = worksheet.getRow(1);
		const choices = firstRow.values.splice(3);
		res.send(choices);
	})
})

app.post('/data',cors(corsOptions), (req, res) => {
	let error = false
	let data = req.body;
	let id = parseInt(data.id);
	let no_of_rows = total;
	id = id * no_of_rows;
	id = id + 2;

	// write to validation file
	validation_workbook.xlsx.readFile(validation_file_path).then(() => {
		const worksheet = validation_workbook.worksheets[0];
		const rows = worksheet.getRows(parseInt(id), no_of_rows)
		const headerRow = worksheet.getRow(1);
		const headers = headerRow.values.slice(1);

		rows.forEach((row, index) => {
			let value = data.values[index]

			for (let i = 1; i < headers.length; i++) {
				if (row.getCell(i + 1).value === '' || row.getCell(i + 1).value === null) {
					row.getCell(i + 1).value = 0;
				}

				if (headers[i].toLowerCase().trim() === value.toLowerCase().trim()) {
					row.getCell(i + 1).value = parseInt(row.getCell(i + 1).value) + 1;
				}
				if (headers[i].toLowerCase().trim() === 'email') {
					if (typeof row.getCell(i + 1).value === 'number')
						row.getCell(i + 1).value = '';

					if (row.getCell(i + 1).value === '' || row.getCell(i + 1).value === null) {
						row.getCell(i + 1).value = data.email.toLowerCase().trim() + ',';
					} else {
						if (!row.getCell(i + 1).value.toLowerCase().trim().includes(data.email.toLowerCase().trim())) {
							row.getCell(i + 1).value += data.email.toLowerCase().trim() + ',';
						} else {
							error = true;
						}
					}
				}
			}
		})

		if (!error)
			validation_workbook.xlsx.writeFile(validation_file_path).then(() => {
				res.status(200).send('success');
				console.log('done');
			})
		else
			return res.status(400).send('email already exists');
	})

})


app.listen(port, () => {
	console.log(`Example app listening at http://localhost:${port}`)
})
