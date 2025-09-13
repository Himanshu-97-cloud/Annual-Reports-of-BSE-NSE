const puppeteerExtra = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const https = require('https');
const XLSX = require('xlsx');
const AdmZip = require('adm-zip');

puppeteerExtra.use(StealthPlugin());

const companyErrors = {}; // { symbol: { name, errors: [] } }

async function addError(symbol, name, message) {
  if (!companyErrors[symbol]) companyErrors[symbol] = { name, errors: [] };
  companyErrors[symbol].errors.push(message);
  console.error(message);
}

function downloadFileWithSpeed(url, dest) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(dest);
    let downloadedBytes = 0;
    let lastDownloadedBytes = 0;

    const request = https.get(url, (response) => {
      if (response.statusCode !== 200) {
        return reject(new Error(`Failed to download '${url}' (Status ${response.statusCode})`));
      }

      const speedInterval = setInterval(() => {
        const speed = downloadedBytes - lastDownloadedBytes;
        lastDownloadedBytes = downloadedBytes;
        const speedInKbps = (speed / 1024).toFixed(2);
        process.stdout.write(`\rDownloading: ${speedInKbps} KB/s`);
      }, 1000);

      response.on('data', (chunk) => {
        downloadedBytes += chunk.length;
      });

      response.pipe(file);

      file.on('finish', () => {
        clearInterval(speedInterval);
        process.stdout.write('\rDownload complete!           \n');
        file.close(resolve);
      });

      response.on('error', (err) => {
        clearInterval(speedInterval);
        reject(err);
      });
    });

    request.on('error', (err) => {
      reject(err);
    });
  });
}

async function extractZipToPdf(zipFilePath, folderPath, symbol, year) {
  try {
    const stats = fs.statSync(zipFilePath);
    if (stats.size < 1024) {
      console.warn(`ZIP file too small (${stats.size} bytes), skipping extraction: ${zipFilePath}`);
      return;
    }

    const zip = new AdmZip(zipFilePath);
    let extracted = false;
    for (const entry of zip.getEntries()) {
      if (entry.entryName.toLowerCase().endsWith('.pdf')) {
        const newPdfName = `${symbol}_${year}.pdf`;
        const outputFilePath = path.join(folderPath, newPdfName);
        fs.writeFileSync(outputFilePath, entry.getData());
        console.log(`Extracted PDF from ZIP and renamed: ${outputFilePath}`);
        extracted = true;
      }
    }
    if (!extracted) console.warn(`No PDF found inside ZIP: ${zipFilePath}`);

    fs.unlinkSync(zipFilePath);
  } catch (err) {
    await addError(symbol, '', `Error extracting ZIP file ${zipFilePath}: ${err.message || err}`);
  }
}

async function scrapeBseReports(page, company) {
  await addError(company.symbol, company.name, `Trying BSE reports for ${company.name} (${company.symbol})`);
  try {
    const baseUrl = 'https://www.bseindia.com/corporates/HistoricalAnnualReport.aspx';
    await page.goto(baseUrl, { waitUntil: 'networkidle2' });
    await page.waitForSelector('#ContentPlaceHolder1_SmartSearch_smartSearch');

    await page.click('#ContentPlaceHolder1_SmartSearch_smartSearch', { clickCount: 3 });
    await page.type('#ContentPlaceHolder1_SmartSearch_smartSearch', company.symbol);

    try {
      await page.waitForSelector('#ulSearchQuote2', { timeout: 10000 });
      await page.click('#ulSearchQuote2 li.quotemenu');
    } catch {
      // ignore autocomplete failure
    }

    await Promise.all([
      page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 20000 }),
      page.click('#ContentPlaceHolder1_btnSubmit'),
    ]);

    await page.waitForSelector('#ContentPlaceHolder1_grdAnnualReport', { timeout: 15000 });

    const folder = path.join(__dirname, 'BSE_AnnualReports', company.symbol);
    fs.mkdirSync(folder, { recursive: true });

    let downloadedAny = false;

    async function extractAndDownloadReports() {
      const rows = await page.$$(`#ContentPlaceHolder1_grdAnnualReport tbody tr`);
      for (let i = 1; i < rows.length; i++) {
        const cells = await rows[i].$$('td');
        if (cells.length < 2) continue;

        const year = await page.evaluate((td) => td.innerText.trim(), cells[0]);
        if (isNaN(parseInt(year, 10)) || parseInt(year, 10) < 2009) continue; // skip before 2009

        const linkHandle = await cells[1].$('a');
        if (!linkHandle) continue;
        const url = await page.evaluate((a) => a.href, linkHandle);
        if (!url.startsWith('https://')) continue;

        let filename = `${company.symbol}_${year}.pdf`;
        filename = filename.replace(/[\\/:"*?<>|]+/g, '').trim();
        const filepath = path.join(folder, filename);

        try {
          await addError(company.symbol, company.name, `Downloading BSE report: ${filename}...`);
          await downloadFileWithSpeed(url, filepath);
          downloadedAny = true;
        } catch (err) {
          await addError(company.symbol, company.name, `Failed to download BSE report ${filename}: ${err.message || err}`);
        }
      }
    }

    await extractAndDownloadReports();

    // Pagination page 2 (optional)
    const nextPageLink = await page.$('a[href*="Page$2"]');
    if (nextPageLink) {
      await Promise.all([
        page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }),
        nextPageLink.click(),
      ]);
      await page.waitForSelector('#ContentPlaceHolder1_grdAnnualReport', { timeout: 10000 });
      await extractAndDownloadReports();
    }

    return downloadedAny;
  } catch (err) {
    await addError(company.symbol, company.name, `BSE scraping failed for ${company.symbol}: ${err.message || err}`);
    return false;
  }
}

async function scrapeNseAnnualReports(browser, symbol, companyName) {
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 1000 });

  try {
    const url = `https://www.nseindia.com/get-quotes/equity?symbol=${symbol}`;
    await page.setUserAgent(
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
    );
    await page.goto(url, { waitUntil: 'networkidle2', timeout: 60000 });

    await page.waitForSelector('#annualReports', { visible: true, timeout: 30000 });
    await page.click('#annualReports');

    await page.waitForSelector('#Annual_Reports', { visible: true, timeout: 30000 });
    await page.waitForSelector('#Annual_Reports table tbody tr', { visible: true, timeout: 30000 });

    const folder = path.join(__dirname, 'NSE_AnnualReports', symbol);
    fs.mkdirSync(folder, { recursive: true });

    const reportRows = await page.$$eval('#Annual_Reports table tbody tr', (rows) =>
      rows
        .map((row) => {
          const yearTd = row.querySelector('td[data-ws-symbol-col*="toYr"]');
          const year = yearTd ? yearTd.textContent.trim() : null;
          const anchor = row.querySelector('td:nth-child(4) a');
          const url = anchor ? anchor.href : null;
          let type = null;
          if (url) {
            const u = url.toLowerCase();
            if (u.endsWith('.zip')) type = 'zip';
            else if (u.endsWith('.pdf')) type = 'pdf';
          }
          return {
            year,
            url,
            type,
          };
        })
        .filter(
          (r) => r.url && r.year && r.type && !isNaN(parseInt(r.year, 10)) && parseInt(r.year, 10) >= 2009
        )
    );

    if (reportRows.length === 0) {
      await page.close();
      return false; // No reports but NOT an error
    }

    for (const report of reportRows) {
      const toYear = report.year;
      let filename = `${symbol}_${toYear}.${report.type}`;
      filename = filename.replace(/[\\/:"*?<>|]+/g, '').trim();
      const destPath = path.join(folder, filename);

      try {
        console.log(`Downloading NSE report: ${filename}...`);
        await downloadFileWithSpeed(report.url, destPath);

        if (report.type === 'zip') {
          await extractZipToPdf(destPath, folder, symbol, toYear);
        }
        if (report.type === 'pdf') {
          const newPdfName = `${symbol}_${toYear}.pdf`;
          const newPdfPath = path.join(folder, newPdfName);
          if (filename !== newPdfName) fs.renameSync(destPath, newPdfPath);
        }
      } catch (e) {
        await addError(symbol, companyName, `Error downloading/extracting NSE report ${filename}: ${e.message || e}`);
      }
    }

    await page.close();
    return true;
  } catch (err) {
    await addError(symbol, companyName, `NSE scraping failed for ${symbol}: ${err.message || err}`);
    try {
      await page.screenshot({ path: `debug_NSE_${symbol}.png`, fullPage: true });
    } catch {}
    await page.close();
    return false;
  }
}

async function main() {
  const browser = await puppeteerExtra.launch({ headless: true });
  const page = await browser.newPage();

  const workbook = XLSX.readFile('companies.xlsx');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet);

  if (data.length === 0) {
    console.error('Excel sheet is empty');
    process.exit(1);
  }

  const allKeys = Object.keys(data[0]);
  const symbolKey = allKeys.find((k) => k.toLowerCase() === 'symbol');
  const companyKey =
    allKeys.find((k) => k.toLowerCase() === 'company name') ||
    allKeys.find((k) => k.toLowerCase() === 'company');

  if (!symbolKey) {
    console.error("No 'symbol' column found in Excel sheet.");
    process.exit(1);
  }

  for (const row of data) {
    const symbol = row[symbolKey] ? String(row[symbolKey]).trim() : null;
    if (!symbol) continue;
    const companyName = companyKey ? row[companyKey] : '';
    const company = { symbol, name: companyName };

    try {
      const bseSuccess = await scrapeBseReports(page, company);
      if (!bseSuccess) {
        const nseSuccess = await scrapeNseAnnualReports(browser, symbol, companyName);
        if (!nseSuccess) {
          await addError(symbol, companyName, `No reports found on BSE or NSE for ${symbol}`);
        }
      }
    } catch (err) {
      await addError(symbol, companyName, `Error processing ${symbol}: ${err.message || err}`);
    }
  }

  if (Object.keys(companyErrors).length > 0) {
    try {
      const rows = Object.entries(companyErrors).map(([symbol, info]) => ({
        Symbol: symbol,
        'Company Name': info.name || '',
        'Error/Status Messages': info.errors.join('\n'),
      }));
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'ErrorsAndMissing');
      XLSX.writeFile(wb, 'not_found_companies.xlsx');
      console.log(`Written ${rows.length} companies with errors or missing reports to 'not_found_companies.xlsx'.`);
    } catch (e) {
      console.error('Failed writing not found companies Excel:', e);
    }
  } else {
    console.log('All companies have reports on BSE or NSE.');
  }

  await browser.close();
  console.log('All done!');
}

main().catch(console.error);


// upated
