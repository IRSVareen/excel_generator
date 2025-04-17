const express = require('express')
const app = express()
const PORT = 3000
const path = require('path')
const config = require('./config.json')

const excellController = require('./controller/excellController')
const jsonData = require('./data')
app.listen(PORT,()=>{
    console.log('Listening...');
})

const txtFilePath = path.resolve(__dirname, config.txtFilePath);
const outputFilePath = path.resolve(__dirname, config.outputFilePath);

excellController.generateExcel(jsonData, outputFilePath);