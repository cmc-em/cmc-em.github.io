/**
 * Simple CSV parser for Google Sheets exports
 */

/**
 * Parses a single CSV line, handling quoted fields
 * @param {string} line - CSV line
 * @returns {string[]} Parsed values
 */
function parseCSVLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === "," && !inQuotes) {
      result.push(current);
      current = "";
    } else {
      current += char;
    }
  }
  result.push(current);
  return result;
}

/**
 * Parses CSV content into array of row objects
 * @param {string} content - CSV file content
 * @returns {Object[]} Array of objects keyed by header names
 */
function parseCSV(content) {
  const lines = content.trim().split("\n");
  const headers = parseCSVLine(lines[0]);
  return lines.slice(1).map((line) => {
    const values = parseCSVLine(line);
    const row = {};
    headers.forEach((h, i) => (row[h.trim()] = values[i]?.trim() || ""));
    return row;
  });
}

module.exports = {
  parseCSV,
  parseCSVLine,
};
