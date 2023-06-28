import express from 'express';
import path, { resolve } from "path";
import { fileURLToPath } from "url";
import fileUpload from "express-fileupload";
import dotenv from "dotenv";
import PDFDocument from "pdfkit";
import fs from "fs";
import mammoth from "mammoth"
import { convertCsvToXlsx } from "@aternus/csv-to-xlsx"

const __filename = fileURLToPath(import.meta.url);
const port = process.env.PORT || 80;
const instance = new express();
const __dirname = path.dirname(__filename);

instance.use(express.static('files'));

dotenv.config();
instance.use(
	fileUpload({
		createParentPath: true,
	})
);

instance.post("/fileupload", (request, response) => {
	if (!request.files) {
		return response.status(400).send("No files are received.");
	}
	const file = request.files.file;
	const path = __dirname + "/files/" + file.name;
	// const path = __dirname + file.name;
	file.mv(path, (err) => {

		if (err) {
			return response.status(500).send(err);
			console.log("Error wysylania")
		}
		console.log(`success ${path}`);
		return response.send({ status: "success", path: path });
	});
});

instance.post("/txt2pdf", async (convert, response) => {
	if (!convert.files) {
		return response.status(400).send("No files are received.");
	} else {
		let outputFileName = await txtPdfConvert(convert.files.file);
		response.download(outputFileName);
	}
});

instance.post("/word2pdf", async (convert, response) => {
	if (!convert.files) {
		return response.status(400).send("No files are received.");
	} else {
		let outputFileName = await wordPdfConvert(convert.files.file);
		response.download(outputFileName);
	}
});

instance.post("/csv2excel", async (convert, response) => {
	if (!convert.files) {
		return response.status(400).send("No files are received.");
	} else {
		let outputFileName = await csvExcelConvert(convert.files.file);
		response.download(outputFileName);
	}
});

function txtPdfConvert(file) {
	return new Promise((resolve, reject) => {
		const nameFile = file.name;
		console.log(`Blob Name ${nameFile}`);

		const outputFilePath = `${nameFile}_output.pdf`;
		fs.unlinkSync(outputFilePath);

		const fileData = file.data;
		const escapedFileData = new String(fileData).toString().replace(/\r\n|\r/g, '\n');
		const writeStream = fs.createWriteStream(outputFilePath, 'UTF-8');
		const pdfDoc = new PDFDocument();
		pdfDoc.pipe(writeStream);
		pdfDoc.text(escapedFileData);
		pdfDoc.end();

		writeStream.on('finish', function () {
			resolve(outputFilePath);
		});
	});
}

function wordPdfConvert(file) {
	return new Promise(async (resolve, reject) => {
		const nameFile = file.name;
		console.log(`Blob Name ${nameFile}`);

		const outputFilePath = `${nameFile}_output.pdf`;
		fs.unlinkSync(outputFilePath);

		try {
			fs.writeFileSync(nameFile, file.data);
			mammoth.extractRawText({ path: nameFile })
				.then(function (result) {
					var textData = result.value;
					const escapedFileData = new String(textData).toString().replace(/\r\n|\r/g, '\n');
					const writeStream = fs.createWriteStream(outputFilePath);
					const pdfDoc = new PDFDocument();
					pdfDoc.pipe(writeStream);
					pdfDoc.text(escapedFileData);
					pdfDoc.end();

					writeStream.on('finish', function () {
						fs.unlinkSync(nameFile);
						resolve(outputFilePath);
					});
				})
				.catch(function (error) {
					console.error(error);
					fs.unlinkSync(nameFile);
					reject(error);
				});
		} catch (err) {
			console.error(err);
			reject(error);
		}

	});
}

function csvExcelConvert(file) {
	return new Promise(async (resolve, reject) => {
		const nameFile = file.name;
		console.log(`Blob Name ${nameFile}`);

		const outputFilePath = `${nameFile}_output.xlsx`;
		fs.unlinkSync(outputFilePath);

		try {
			fs.writeFileSync(nameFile, file.data);

			try {
				convertCsvToXlsx(nameFile, outputFilePath);

				fs.unlinkSync(nameFile);
				resolve(outputFilePath);
			} catch (e) {
				console.error(e.toString());
				reject(e);
			}

			fs.unlinkSync(nameFile);

		} catch (err) {
			console.error(err);
			reject(err);
		}

	});
}

instance.listen(port, () => {
	console.log(`Server is listening on port ${port}`);
});

instance.get('/', (req, res) => {
	res.sendFile(path.join(__dirname, "index.html"));
})

instance.get('/txt2pdf', (req, res) => {
	res.sendFile(path.join(__dirname, "txt2pdf.html"));
})

instance.get('/word2pdf', (req, res) => {
	res.sendFile(path.join(__dirname, "word2pdf.html"));
})

instance.get('/csv2excel', (req, res) => {
	res.sendFile(path.join(__dirname, "csv2excel.html"));
})