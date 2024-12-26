class KeepaScrape {

    constructor(apiKey) {

        this.apiKey = apiKey;
        this.baseUri = 'https://api.keepa.com/';

    }

    getProduct(asinList, domainid) {
        const uri = `${this.baseUri}product?key=${this.apiKey}&domain=${domainid}&asin=${asinList.join(',')}&stats=1&buybox=1&stock=1&offers=20`;
        Logger.log(`Calling Keepa API with URI: ${uri}`);
        try {
            const resp = this.callAPI(uri);
            return resp.products || [];
        } catch (error) {
            Logger.log(`Failed to fetch products for ASINs: ${asinList.join(', ')}. Error: ${error.message}`);
            return [];
        }
    }

    getTokenLeft() {
        const uri = `${this.baseUri}token?key=${this.apiKey}`;
        const resp = this.callAPI(uri);
        return resp.tokensLeft;
    }

    callAPI(uri) {
        try {
            const response = UrlFetchApp.fetch(encodeURI(uri));
            if (response.getResponseCode() !== 200) {
                Logger.log(`Error: ${response.getResponseCode()} - ${response.getContentText()}`);
                throw new Error(`Request failed with response code ${response.getResponseCode()}`);
            }
            return JSON.parse(response.getContentText());
        } catch (e) {
            Logger.log(`API call failed: ${e.message}`);
            throw e;
        }
    }

    static getWeight(product) {
        return product.weight ? product.weight / 100 : null;
    }

    static getReferralFee(product) {
        return product.referralFeePercentage ? product.referralFeePercentage.toFixed(2) : '';
    }

    static getBuyboxPrice(price) {
        return price ? `${(price / 100.0).toFixed(2)}` : '';
    }
    static fbaSellerCount(product) {
    var totalFBASellerCount = 0;
    let offers = 0;
    if (product.offers != null) {
      offers = product.offers.length;

      for (var i = 0; i < offers; i++) {
        var Offer = product.offers[i];
        if (Offer.isFBA)
          totalFBASellerCount++;
      }
    }
    return totalFBASellerCount;
  }
  static livefbaSellerCount(product) {
    var totalFBASellerCount = 0;
    let totalLiveOfferCount = 0;
    if (product.liveOffersOrder != null) {
      totalLiveOfferCount = product.liveOffersOrder.length;

      for (var i = 0; i < totalLiveOfferCount; i++) {
        var offerOrderNo = product.liveOffersOrder[i];
        var Offer = product.offers[offerOrderNo];
        if (Offer.isFBA)
          totalFBASellerCount++;
      }
    }
    return totalFBASellerCount;
  }
}

class ASINFetcher {
    constructor(spreadsheetId, apiUrl, username, password, keepaApiKey) {
        this.spreadsheetId = spreadsheetId;
        this.apiUrl = apiUrl;
        this.username = username;
        this.password = password;
        this.ss = SpreadsheetApp.openById(this.spreadsheetId);
        this.keepaScrapes = new KeepaScrape(keepaApiKey);
        this.batchSize = 60;
        this.lastProcessedCells = 'A1';
    }

    main() {
        const dataSheet = this.ss.getSheetByName("OALive"); 
        const setupSheet = this.ss.getSheetByName("OALive"); 
        let setup = this.ss.getSheetByName('Setup');
        let exchange_rate = setup.getRange('B10').getValue();
        let weight_factor = setup.getRange('B9').getValue();
        let prep_fees = setup.getRange('B8').getValue();

        const username = this.username;
        const password = this.password;
        const apiToken = Utilities.base64Encode(username + ":" + password); 

        const d2Value = setupSheet.getRange('A1').getValue();
        const domain_ID = parseInt(d2Value.match(/\d+/)[0], 10);

        let lastProcessedIndex = parseInt(dataSheet.getRange(this.lastProcessedCells).getValue(), 10) || 2;
        const combinedData = dataSheet.getRange(
            lastProcessedIndex + 1, 
            1,
            dataSheet.getLastRow() - lastProcessedIndex,
            2
        ).getValues();

        if (combinedData.length === 0) {
            Logger.log('No data left to process.');
            return;
        }

        const batch = combinedData.slice(0, this.batchSize);
        Logger.log(`Processing ${batch.length} ASINs starting from index ${lastProcessedIndex + 1}`);

        batch.forEach((item, index) => {
            const asin = item[0];
            const url = item[1];
            const row = lastProcessedIndex + 1 + index; // Calculate row in sheet

            if (asin !== 'No Results') {
                const productList = this.keepaScrapes.getProduct([asin], domain_ID);
                const output = RequiredOutput(productList);
                if (output.length > 0) {
                    const tookn = this.keepaScrapes.getTokenLeft()
                    const [title, bbPrice,pastMonthSold, srdrp30, weight, referralFeePercentage, fbaFee, stock,saturationScore, avgBuyBox30,avgBuyBox90,avgBuyBox180] = output[0];
             
                    const bbPriceNumeric = (bbPrice && typeof bbPrice === "string") ? parseFloat(bbPrice.replace('$', '')) : (typeof bbPrice === "number" ? bbPrice : 0);
                    const fbaFeeNumeric = typeof fbaFee === 'string' ? parseFloat(fbaFee.replace('$', '').trim()) : parseFloat(fbaFee) || 0;
                    const referralFee = (referralFeePercentage / 100) * bbPriceNumeric;
                    const refFee = referralFee.toFixed(2);
                    const apiResponse = this.fetchOxylabsData(url, apiToken);
                    if (apiResponse) {
                        const price = apiResponse.price || "-";
                        const availability = apiResponse.availability ? "TRUE" : "FALSE";
                        const storePriceStr = String(price || price || '0');
                        const storePriceGBP = parseFloat(storePriceStr.replace(/[^0-9.]/g, ''));
                        const euroToUsd = storePriceGBP * exchange_rate;
                        const profit = bbPriceNumeric - euroToUsd - 1.04 - (weight * 0.01) - referralFee - fbaFeeNumeric;
                        const roi = (profit / (euroToUsd + 1.04 + (weight * 0.01))) * 100;
                        const formattedTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

                        let mean_price = (avgBuyBox30 + avgBuyBox90 + avgBuyBox180) / 3;
                        let mean_price_rounded = parseFloat(mean_price.toFixed(2));
                        let bbprices = [avgBuyBox30, avgBuyBox90, avgBuyBox180];
                        if (bbprices.length === 0) return 0;
                        let mean = bbprices.reduce((sum, value) => sum + value, 0) / bbprices.length;
                        let variance = bbprices.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) / bbprices.length;
                        let standardDeviation = Math.sqrt(variance).toFixed(2);
                        let upper_limit_price = mean_price_rounded + (2 * standardDeviation);
                        let lower_limit_price = mean_price_rounded - (2 * standardDeviation);
                      
                        let upper_limit_roi =(upper_limit_price - refFee - fbaFee - (price*exchange_rate)-(weight-weight_factor)-prep_fees)/(price*exchange_rate); 
                        let lower_limit_roi = (lower_limit_price - refFee - fbaFee - (price*exchange_rate)-(weight-weight_factor)-prep_fees)/(price*exchange_rate);
                        let upper_limit_roi_rounded = parseFloat(upper_limit_roi.toFixed(2));
                        let lower_limit_roi_rounded = parseFloat(lower_limit_roi.toFixed(2));
                        let quantity = pastMonthSold - stock;
                        quantity = quantity > 0 ? quantity : 0;
                        let purchaseDecision = false;
                        if (saturationScore < 1 && roi > 30 && availability === 'TRUE') {
                            if (lower_limit_roi_rounded > 20) {
                                purchaseDecision = true; // All conditions for purchase are met
                            }
                        }
                        let adjustedPurchaseAmount = quantity;
                        if (lower_limit_roi_rounded < 20 && purchaseDecision ==='False') {
                            let scaleFactor = 0.25;
                            adjustedPurchaseAmount = quantity * scaleFactor;
                        }
                        let lower_limit_roi_probability = `=ROUND(NORM.DIST(${lower_limit_price}, ${mean_price}, ${standardDeviation}, TRUE) * 100, 2)`;
                        let upper_limit_roi_probability = `=ROUND(NORM.DIST(${upper_limit_price}, ${mean_price}, ${standardDeviation}, TRUE) * 100, 2)`;
                        let round_stauration = saturationScore.toFixed(2);
                        dataSheet.getRange(row, 3).setValue(price);
                        dataSheet.getRange(row, 4).setValue(profit.toFixed(2));
                        dataSheet.getRange(row, 5).setValue(roi.toFixed(2) + '%');
                        dataSheet.getRange(row, 6).setValue(availability);
                        dataSheet.getRange(row, 7).setValue(formattedTimestamp);
                        dataSheet.getRange(row, 10).setValue(title);
                        dataSheet.getRange(row, 11).setValue(`$${bbPriceNumeric.toFixed(2)}`);
                        dataSheet.getRange(row, 12).setValue(pastMonthSold);
                        dataSheet.getRange(row, 13).setValue(srdrp30);
                        dataSheet.getRange(row, 14).setValue(weight || 0);
                        dataSheet.getRange(row, 15).setValue(referralFeePercentage?.toFixed(2) || 'N/A');
                        dataSheet.getRange(row, 16).setValue(`$${refFee}`);
                        dataSheet.getRange(row, 17).setValue(fbaFeeNumeric);
                        dataSheet.getRange(row, 18).setValue(stock);
                        dataSheet.getRange(row, 19).setValue(round_stauration);
                        dataSheet.getRange(row, 20).setValue(purchaseDecision);
                        dataSheet.getRange(row, 21).setValue(adjustedPurchaseAmount);
                        dataSheet.getRange(row, 22).setValue(avgBuyBox30);
                        dataSheet.getRange(row, 23).setValue(avgBuyBox90);
                        dataSheet.getRange(row, 24).setValue(avgBuyBox180);
                        dataSheet.getRange(row, 25).setValue(mean_price);
                        dataSheet.getRange(row, 26).setValue(standardDeviation);  
                        dataSheet.getRange(row, 27).setValue(upper_limit_price); 
                        dataSheet.getRange(row, 28).setValue(lower_limit_price); 
                        dataSheet.getRange(row, 29).setValue(lower_limit_roi_rounded); 
                        dataSheet.getRange(row, 30).setValue(upper_limit_roi_rounded); 
                        dataSheet.getRange(row, 31).setFormula(lower_limit_roi_probability);
                        dataSheet.getRange(row, 32).setFormula(upper_limit_roi_probability);
                        Logger.log(`Sheet Row ${row}: Lower Limit Probability Formula = ${lower_limit_roi_probability}`);
                        Logger.log(`Sheet Row ${row}: Upper Limit Probability Formula = ${upper_limit_roi_probability}`);

                        
                    }
                    Logger.log(`token left ${tookn}`);
                    
                    Logger.log(`Wrote entry to report sheet at row ${row}`);
                }
            } else {
                Logger.log(`Skipping entry with ASIN 'No Results' for: `);
            }
        });

        const newLastProcessedIndex = lastProcessedIndex + batch.length;
        // dataSheet.getRange(this.lastProcessedCells).setValue(newLastProcessedIndex);
        Logger.log(`Updated last processed index to ${newLastProcessedIndex}`);
    }


    fetchOxylabsData(url, apiToken) {
        const apiUrl = "https://realtime.oxylabs.io/v1/queries";
        const payload = {
            source: "universal",
            url: url,
            parse: true
        };

        const options = {
            method: "post",
            contentType: "application/json",
            headers: {
                Authorization: "Basic " + apiToken
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        try {
            const response = UrlFetchApp.fetch(apiUrl, options);
            const statusCode = response.getResponseCode();

            if (statusCode !== 200) {
                Logger.log("Error Response: " + response.getContentText());
                return null;
            }

            const json = JSON.parse(response.getContentText());
            const content = json.results[0]?.content;

            return {
                price: content.price,
                availability: content.availability
            };
        } catch (error) {
            Logger.log("Error fetching data: " + error.message);
            return null;
        }
    }
}

function RequiredOutput(products) {
    return products.map(product => {
        let title = product.title || '-';
        let pastMonthSold = product.monthlySold;
        let bbprice = KeepaScrape.getBuyboxPrice(product.stats.buyBoxPrice) || '0';
        let srdrp30 = product.stats.salesRankDrops30 || 0;
        let referralFeePercentage = product.referralFeePercentage || 0;
        let weight = product.packageWeight || 0;
        let stockArray = product.stats.stockPerCondition3rdFBA || [];
        let stock = stockArray[1] || 0;
        let avg30 = product.stats.avg30 || [];
        let avg90 = product.stats.avg90 || [];
        let avg180 = product.stats.avg180 || [];
        let avgBuyBox30 = parseFloat(KeepaScrape.getBuyboxPrice(avg30[1])) || 0;
        let avgBuyBox90 = parseFloat(KeepaScrape.getBuyboxPrice(avg90[1])) || 0;
        let avgBuyBox180 = parseFloat(KeepaScrape.getBuyboxPrice(avg180[1])) || 0;
        let fbaFee = '-';
        if (product.fbaFees && product.fbaFees.pickAndPackFee != null) {fbaFee = KeepaScrape.getBuyboxPrice(product.fbaFees.pickAndPackFee);}
        let saturationScore = stock/pastMonthSold;
        
        return [
            title, bbprice, pastMonthSold, srdrp30, weight, referralFeePercentage, fbaFee,
            stock, saturationScore, avgBuyBox30, avgBuyBox90, avgBuyBox180
        ];
    });
}

function runGoogleScraper() {
    const sheetId = '1Ro8ELkH-IC5oBEXZ_lD_OH-JQlrSee_sV_DBA78k9oA';
    const sheetName = 'Setup';
    const credentials = getCredentialsFromSheet(sheetId, sheetName);
    const fetcher = new ASINFetcher(
        sheetId,
        credentials.apiUrl,
        credentials.username,
        credentials.password,
        credentials.keepaApiKey
    );

    fetcher.main();
}

function getCredentialsFromSheet(sheetId, sheetName) {
    try {
        const ss = SpreadsheetApp.openById(sheetId);
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) throw new Error('Sheet not found: ' + sheetName);
        const data = sheet.getDataRange().getValues();
        const credentials = {         
            apiUrl: '',
            username: '',
            password: '',
            keepaApiKey: '' 
        };
        data.forEach(row => {
            if (row[0] === 'apiUrl') credentials.apiUrl = row[1];
            if (row[0] === 'username') credentials.username = row[1];
            if (row[0] === 'password') credentials.password = row[1];
            if (row[0] === 'keepaApiKey') credentials.keepaApiKey = row[1];
        });
        return credentials;
    } catch (error) {
        Logger.log('Error getting credentials: ' + error);
        throw error;
    }
}
