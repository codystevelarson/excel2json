import * as XLSX from 'xlsx';
import { parse } from 'csv-parse';
import path from 'path';
import fs from 'fs';

const [, , input, output, arrayKeysArg] = process.argv;
const ARRAY_KEYS = arrayKeysArg ? arrayKeysArg.split(',').map(k => k.trim()) : [];

const SUPPORTED_TYPES = ['csv', 'xlsx', 'xls'];

export function excelToJson(filePath) {
    // Read the workbook
    const workbook = XLSX.default.readFileSync(filePath);

    // Use the first sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert to JSON
    const jsonData = XLSX.utils
        .sheet_to_json(sheet, { defval: "" })
        .map(mapRowKeysToCamelCase)
        .map(normalizeRow);

    return jsonData;
}

export async function csvToJson(filePath) {
    const fileContent = fs.readFileSync(filePath, 'utf8');

    const records = [];
    const parser = parse(fileContent, {
        columns: header => header.map(toCamelCase),
        on_record: record => normalizeRow(record),
        skip_empty_lines: true
    });

    for await (const record of parser) {
        records.push(record);
    }

    return records;
}

export function jsonToFile(jsonData, outputPath) {
    if (!outputPath) {
        console.log(jsonData);
        return;
    }

    const dir = path.dirname(outputPath);
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }

    fs.writeFileSync(`${outputPath}-${format}.json`, JSON.stringify(jsonData, null, 2), "utf8");

    return outputPath;
}

function normalizeRow(row) {
    return Object.fromEntries(
        Object.entries(row).map(([key, value]) => {
            if (typeof value === "string") value = value.trim();
            let val = parseValue(value);
            if (!ARRAY_KEYS.includes(key)) return [key, val];

            let vals = typeof val === "string" 
                ? val.split(",").map(x => x.trim())
                : [val];
                return [key, vals];
        })
    );
}

function parseValue(value) {
    const str = String(value).trim();
    if (/^[+-]?\d+$/.test(str)) {
        return Number(str);
    }

    if (/^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(str)) {
        return Number(str);
    }

    if (str === "true" || str === "false") {
        return Boolean(str)
    }

    return value;
}

function toCamelCase(str) {
    return str
        .replace(/[^a-zA-Z0-9 ]/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .split(" ")
        .map((word, index) => index === 0 ? word.toLowerCase() : word[0].toUpperCase() + word.slice(1).toLowerCase())
        .join("");
}

function mapRowKeysToCamelCase(row) {
    return Object.fromEntries(
        Object.entries(row).map(([key, value]) => [toCamelCase(key), value])
    );
}

// CLI
if (!input) {
    console.error("Missing input path");
    process.exit(1);
}

const format = input.split('.').at(-1);
if (!SUPPORTED_TYPES.includes(format)) {
    console.error(`Invalid input file format - supported: ${SUPPORTED_TYPES.join(', ')}`);
    process.exit(1);
}

if (!output) {
    console.warn('No output specified, writing to console');
}

const json = input.endsWith('.csv') ? await csvToJson(input) : excelToJson(input);
jsonToFile(json, output);
