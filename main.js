
// INSTRUCTIONS FOR USE

// copy the template sheet & then grab YOUR url & paste that below in SHEET_URL
// template is: 
//
//       https://docs.google.com/spreadsheets/d/16P6WveKp6jQ7_YVLR_XOKJdy55JZeLCc69yflcV3GKA/copy


// then make changes to the variables below as needed, or if you prefer use the 'Settings' tab in your sheet


  let SHEET_URL    = 'https://docs.google.com/spreadsheets/d/1Va8mD9Fn1gC6z8Dv0vsxwhjcM56D0J6syscw4zBiVq4/'      // create a copy of the template first & enter YOUR URL between the single quotes
  let clientCode   = ''      // this string will be added to the sheet name use it to easily identify each account/sheet
  let numberOfDays = 30      // how many days data do you want to get?    
  let tCost        = 10      // the 'threshold' for low vs high cost - see the 'map' on Product tab in sheet
  let tRoas        = 4       // the 'threshold' for low vs high roas (performance)    
  let productDays  = 90      // how many days data do you want to use to create product 'buckets'?    
  let brandTerm    = ''      // a string that best represents the brand name - used to find the search category 'buckets'
  




  
  function main() {
    
// Setup Sheet
    let accountName = AdsApp.currentAccount().getName();
    let sheetNameIdentifier = clientCode ? clientCode : accountName;
    let ss = SpreadsheetApp.openByUrl(SHEET_URL);
    ss.rename(sheetNameIdentifier + ' JP - PMAX Insights ');
    updateVariablesFromSheet(ss, 'Settings');


// define query elements. wrap with spaces for safety
    let impr       = ' metrics.impressions ';
    let clicks     = ' metrics.clicks ';
    let cost       = ' metrics.cost_micros ';
    let engage     = ' metrics.engagements ';
    let inter      = ' metrics.interactions ';
    let conv       = ' metrics.conversions ';
    let value      = ' metrics.conversions_value ';
    let allConv    = ' metrics.all_conversions ';
    let allValue   = ' metrics.all_conversions_value ';
    let views      = ' metrics.video_views ';
    let cpv        = ' metrics.average_cpv ';
    let segDate    = ' segments.date ';
    let prodTitle  = ' segments.product_title ';
    let prodID     = ' segments.product_item_id ';
    let campName   = ' campaign.name ';
    let campId     = ' campaign.id ';
    let chType     = ' campaign.advertising_channel_type ';
    let aIdAsset   = ' asset.resource_name ';
    let aId        = ' asset.id ';
    let assetSource= ' asset.source ';
    let agId       = ' asset_group.id ';
    let assetFtype = ' asset_group_asset.field_type ';
    let adPmaxPerf = ' asset_group_asset.performance_label ';
    let agStrength = ' asset_group.ad_strength ';
    let agStatus   = ' asset_group.status ';
    let asgName    = ' asset_group.name ';
    let lgType     = ' asset_group_listing_group_filter.type ';
    let aIdCamp    = ' segments.asset_interaction_target.asset ';    
    let assetName  = ' asset.name ';
    let adUrl      = ' asset.image_asset.full_size.url ';
    let ytTitle    = ' asset.youtube_video_asset.youtube_video_title ';
    let ytId       = ' asset.youtube_video_asset.youtube_video_id '; 
    let interAsset = ' segments.asset_interaction_target.interaction_on_this_asset ';   
    let pMaxOnly   = ' AND campaign.advertising_channel_type = "PERFORMANCE_MAX" '; 
    let agFilter   = ' AND asset_group_listing_group_filter.type != "SUBDIVISION" ';
    let notInter   = ' AND segments.asset_interaction_target.interaction_on_this_asset != "TRUE" ';
    let order      = ' ORDER BY campaign.name ';
    

    
// Date stuff --------------------------------------------
// Get the time zone of the current Google Ads account
    let timeZone = AdsApp.currentAccount().getTimeZone();
// Current date ranges
    let today = new Date();
    let yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);              // Yesterday
    let startDate = new Date(today);
    startDate.setDate(today.getDate() - numberOfDays);   // Start date based on numberOfDays from config
    let productStart = new Date(today);
    productStart.setDate(today.getDate() - productDays); // Start date for long-range products from config
// Format Dates
    let formattedYesterday = Utilities.formatDate(yesterday, timeZone, 'yyyy-MM-dd');
    let formattedStart     = Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd');
    let formattedProduct   = Utilities.formatDate(productStart, timeZone, 'yyyy-MM-dd');
// SQL Date Ranges
    let mainDateRange      = ` segments.date BETWEEN "${formattedStart}" AND "${formattedYesterday}" `; 
    let prodDateRange      = ` segments.date BETWEEN "${formattedProduct}" AND "${formattedYesterday}" `;

    
// Build queries --------------------------------------------
    let assetGroupAssetColumns = [campName, asgName, agId, aIdAsset, assetFtype, campId]; //  adPmaxPerf, agStrength, agStatus, assetSource,
    let assetGroupAssetQuery = 'SELECT ' + assetGroupAssetColumns.join(',') +
      ' FROM asset_group_asset ' + 
      ' WHERE campaign.status != REMOVED ' + pMaxOnly ;

    let displayVideoColumns = [segDate, campName, aIdCamp, cost, conv, value, views, cpv, impr, clicks, chType, interAsset, campId];
    let displayVideoQuery = 'SELECT ' + displayVideoColumns.join(',') +
      ' FROM campaign ' +
      ' WHERE ' + mainDateRange + pMaxOnly + notInter + order;

    let assetGroupColumns = [segDate, campName, asgName, agStrength, agStatus, lgType, impr, clicks, cost, conv, value]; 
    let assetGroupQuery = 'SELECT ' + assetGroupColumns.join(',')  + 
      ' FROM asset_group_product_group_view ' +
      ' WHERE ' + mainDateRange + agFilter ;

    let campaignColumns = [segDate, campName, cost, conv, value, views, cpv, impr, clicks, chType, campId]; 
    let campaignQuery = 'SELECT ' + campaignColumns.join(',') + 
      ' FROM campaign ' +
      ' WHERE ' + mainDateRange + pMaxOnly  + order;

    let productColumns = [prodTitle, prodID, cost, conv, value, impr, clicks, chType]; 
    let productQuery = 'SELECT ' + productColumns.join(',')  + 
      ' FROM shopping_performance_view  ' + 
      ' WHERE metrics.impressions > 0 AND ' + prodDateRange + pMaxOnly ; 

    let assetColumns = [aIdAsset, assetSource, ytTitle, ytId, assetName] 
    let assetQuery = 'SELECT ' + assetColumns.join(',')  + 
      ' FROM asset ' ;

    
// Process data --------------------------------------------    

    let assetGroupAssetData = fetchData(assetGroupAssetQuery);
    let displayVideoData    = fetchData(displayVideoQuery);
    let assetGroupData      = fetchData(assetGroupQuery);
    let campaignData        = fetchData(campaignQuery);
    let assetData           = fetchData(assetQuery); 
    let productData         = fetchProductData(productQuery, 'products');  
    let productData2        = fetchProductData(productQuery, 'prod2');   
    let productDataID       = fetchProductData(productQuery, 'id');   

    
// Extract marketing assets & de-dupe
    let { displayAssets, videoAssets } = extractAndDeduplicateAssets(assetGroupAssetData);

    
// Filter displayVideoData based on the type
    let displayAssetData = filterDataByAssets(displayVideoData, displayAssets);
    let videoAssetData   = filterDataByAssets(displayVideoData, videoAssets);

    
// Process data for each dataset
    let processedDisplayAssetData = aggregateDataByDateAndCampaign(displayAssetData);
    let processedVideoAssetData   = aggregateDataByDateAndCampaign(videoAssetData);  
    let processedAssetGroupData   = processData(assetGroupData);
    let processedCampData         = processData(campaignData);
  
// Combine all non-search metrics, calc 'search' & process summary
    let nonSearchData = [...processedDisplayAssetData, ...processedVideoAssetData, ...processedAssetGroupData];
    let searchResults = getSearchResults(processedCampData, nonSearchData);
    let totalData     = processTotalData(processedCampData, processedAssetGroupData, processedDisplayAssetData, processedVideoAssetData, searchResults);
    let summaryData   = processSummaryData(processedCampData, processedAssetGroupData, processedDisplayAssetData, processedVideoAssetData, searchResults);    


// Aggregate the metrics for display and video assets
    let aggregatedDisplayAssetMetrics = aggregateMetricsByAsset(displayAssetData);
    let aggregatedVideoAssetMetrics   = aggregateMetricsByAsset(videoAssetData);
    let enrichedDisplayAssetDetails   = enrichAssetMetrics(aggregatedDisplayAssetMetrics, assetData, 'display');
    let enrichedVideoAssetDetails     = enrichAssetMetrics(aggregatedVideoAssetMetrics, assetData, 'video');
    let mergedDisplayData             = mergeMetricsWithDetails(aggregatedDisplayAssetMetrics, enrichedDisplayAssetDetails);
    let mergedVideoData               = mergeMetricsWithDetails(aggregatedVideoAssetMetrics, enrichedVideoAssetDetails);

    
// Output the data to respective sheets
    outputAggregatedDataToSheet(ss, 'display',  mergedDisplayData);
    outputAggregatedDataToSheet(ss, 'video',    mergedVideoData);
    outputSummaryToSheet(       ss, 'totals',   totalData);
    outputSummaryToSheet(       ss, 'summary',  summaryData);
    outputDataToSheet(          ss, 'group',    assetGroupData,[0,8,12,13,14]);
    outputDataToSheet(          ss, 'products', productData);
    outputDataToSheet(          ss, 'prod2',    productData2);
    outputDataToSheet(          ss, 'id',       productDataID);
    
// get search terms for main date range  
    let levenshteinTolerance = 2 // only change this if you really need to, to change the search category buckets
    extractTerms(ss, mainDateRange, levenshteinTolerance);
    // extractNGrams(ss, mainDateRange); // not used yet
    
    
} // end main




// various additional functions ---------------------------------------------




// fetch data given a query string using search (not report)
  function fetchData(queryString) {
    let data = [];
    const iterator = AdsApp.search(queryString);
    while (iterator.hasNext()) {
      const row = iterator.next();
      const rowData = flattenObject(row); // Flatten the row data
      data.push(rowData);
    }
    return data;
  }




// flatten the object data to enable more processing of that data
  function flattenObject(ob) {
    var toReturn = {};
    for (var i in ob) {
      if ((typeof ob[i]) === 'object') {
        var flatObject = flattenObject(ob[i]);
        for (var x in flatObject) {
          toReturn[i + '.' + x] = flatObject[x];
        }
      } else {
        toReturn[i] = ob[i];
      }
    }
    return toReturn;
  }



// extract & dedupe in one step
  function extractAndDeduplicateAssets(data) {
      let displayAssetsSet = new Set();
      let videoAssetsSet = new Set();

      data.forEach(row => {
          if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('MARKETING')) {
              displayAssetsSet.add(row['asset.resourceName']);
          }
          if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('VIDEO')) {
              videoAssetsSet.add(row['asset.resourceName']);
          }
      });

      return {
          displayAssets: [...displayAssetsSet],
          videoAssets: [...videoAssetsSet]
      };
  }



// filter the display & video assets
  function filterDataByAssets(data, assets) {
      return data.filter(row => assets.includes(row['segments.assetInteractionTarget.asset']));
  }



// get additional data from asset for video 
  function mergeMetricsWithDetails(aggregatedVideoAssetMetrics, enrichedVideoAssetDetails) {
      return enrichedVideoAssetDetails.map(detail => {
          const metrics = aggregatedVideoAssetMetrics[detail.assetName];
          return {
              ...detail, 
              ...metrics
          };
      });
  }



// enrich assets
  function enrichAssetMetrics(aggregatedMetrics, assetData, type) {
      let assetDetailsArray = [];

      // For each asset in aggregatedMetrics, fetch details from assetData
      for (let assetName of Object.keys(aggregatedMetrics)) {
          // Find the asset in assetData
          let matchingAsset = assetData.find(asset => asset['asset.resourceName'] === assetName);

          if (matchingAsset) {
              let assetDetails = {
                  assetName: assetName,
                  assetSource: matchingAsset['asset.source'],
                  assetImage: matchingAsset['asset.name'] 
              };

              if (type === 'video') {
                  assetDetails.youtubeTitle = matchingAsset['asset.youtubeVideoAsset.youtubeVideoTitle'];
                  assetDetails.youtubeId = matchingAsset['asset.youtubeVideoAsset.youtubeVideoId'];
              }

              assetDetailsArray.push(assetDetails);
          }
      }

      return assetDetailsArray;
  }



// Aggregate display & video assets to show total metrics for each asset in their sheets
  function aggregateMetricsByAsset(data) {
      const aggregatedData = {};

      data.forEach(row => {
          const asset = row['segments.assetInteractionTarget.asset'];

          if (!aggregatedData[asset]) {
              aggregatedData[asset] = {
                  clicks: 0,
                  videoViews: 0,
                  conversionsValue: 0,
                  conversions: 0,
                  cost: 0,
                  impressions: 0
              };
          }

          aggregatedData[asset].clicks += parseInt(row['metrics.clicks']);
          aggregatedData[asset].videoViews += parseInt(row['metrics.videoViews']);
          aggregatedData[asset].conversionsValue += parseFloat(row['metrics.conversionsValue']);
          aggregatedData[asset].conversions += parseInt(row['metrics.conversions']);
          aggregatedData[asset].cost += parseInt(row['metrics.costMicros']) / 1e6; 
          aggregatedData[asset].impressions += parseInt(row['metrics.impressions']);
      });

      return aggregatedData;
  }



// function to output the display & video metrics to tabs
  function outputAggregatedDataToSheet(ss, sheetName, data) {
      // Create or access the sheet
      let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

      // Clear the sheet
      sheet.clear();

      // Set the header based on sheetName
      const headers = sheetName === 'video' 
          ? ['Asset Name', 'Source', 'YouTube Title', 'YouTube ID', 'Impr', 'Clicks', 'Views', 'Value', 'Conv', 'Cost']
          : ['Asset Name', 'Source', 'ImageName', 'Impr', 'Clicks', 'Views', 'Value', 'Conv', 'Cost'];
      sheet.appendRow(headers);

      // Append the aggregated data
      data.forEach(item => {
          const rowData = sheetName === 'video'
              ? [item.assetName, item.assetSource, item.youtubeTitle, item.youtubeId, item.impressions, item.clicks, item.videoViews, item.conversionsValue, item.conversions, item.cost]
              : [item.assetName, item.assetSource, item.assetImage, item.impressions, item.clicks, item.videoViews, item.conversionsValue, item.conversions, item.cost];
          sheet.appendRow(rowData);
      });
  }



// Process data: Aggregate by date, campaign name, and various metrics
  function aggregateDataByDateAndCampaign(data) {
      const aggregatedData = {};

      data.forEach(row => {
          const date = row['segments.date'];
          const campaignName = row['campaign.name'];
          const key = `${date}_${campaignName}`;

          if (!aggregatedData[key]) {
              aggregatedData[key] = {
                  'date': date,
                  'campaignName': campaignName,
                  'cost': 0,
                  'impressions': 0,
                  'clicks': 0,
                  'conversions': 0,
                  'conversionsValue': 0
              };
          }

          if (row['metrics.costMicros']) {
              aggregatedData[key].cost += row['metrics.costMicros'] / 1000000;
          }
          if (row['metrics.impressions']) {
              aggregatedData[key].impressions += parseInt(row['metrics.impressions']);
          }
          if (row['metrics.clicks']) {
              aggregatedData[key].clicks += parseInt(row['metrics.clicks']);
          }
          if (row['metrics.conversions']) {
              aggregatedData[key].conversions += parseFloat(row['metrics.conversions']);
          }
          if (row['metrics.conversionsValue']) {
              aggregatedData[key].conversionsValue += parseFloat(row['metrics.conversionsValue']);
          }
      });

      return Object.values(aggregatedData);
  }



// calc the search results
  function getSearchResults(processedCampData, nonSearchData) {
    // Initialize an object to hold 'search' metrics
    const searchMetrics = {};

    // Sum up the non-search metrics
    nonSearchData.forEach(row => {
      const key = `${row.date}_${row.campaignName}`;
      if (!searchMetrics[key]) {
        searchMetrics[key] = {
          date: row.date,
          campaignName: row.campaignName,
          cost: 0,
          impressions: 0,
          clicks: 0,
          conversions: 0,
          conversionsValue: 0
        };
      }
      for (let metric of ['cost', 'impressions', 'clicks', 'conversions', 'conversionsValue']) {
        searchMetrics[key][metric] += row[metric];
      }
    });

    // Calculate 'search' metrics
    const searchResults = [];
    processedCampData.forEach(row => {
      const rowCopy = JSON.parse(JSON.stringify(row));  // Deep copy
      const key = `${row.date}_${row.campaignName}`;
      if (searchMetrics[key]) {
        for (let metric of ['cost', 'impressions', 'clicks', 'conversions', 'conversionsValue']) {
          rowCopy[metric] -= (searchMetrics[key][metric] || 0);
        }
      }
      searchResults.push(rowCopy);
    });

    return searchResults;
  }



// process the data
  function processData(data) {
    const summedData = {};
    data.forEach(row => {
      const date = row['segments.date'];
      const campaignName = row['campaign.name'];
      const key = `${date}_${campaignName}`;
      // Initialize if the key doesn't exist
      if (!summedData[key]) {
        summedData[key] = {
          'date': date,
          'campaignName': campaignName,
          'cost': 0,
          'impressions': 0,
          'clicks': 0,
          'conversions': 0,
          'conversionsValue': 0
        };
      }
      if (row['metrics.costMicros']) {
        summedData[key].cost += row['metrics.costMicros'] / 1000000;
      }
      if (row['metrics.impressions']) {
        summedData[key].impressions += parseInt(row['metrics.impressions']);
      }
      if (row['metrics.clicks']) {
        summedData[key].clicks += parseInt(row['metrics.clicks']);
      }
      if (row['metrics.conversions']) {
        summedData[key].conversions += parseFloat(row['metrics.conversions']);
      }
      if (row['metrics.conversionsValue']) {
        summedData[key].conversionsValue += parseFloat(row['metrics.conversionsValue']);
      }
    });

    return Object.values(summedData);
  }



// output data to tab in sheet
function outputDataToSheet(spreadsheet, sheetName, data, indexesToRemove) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
    }
    sheet.clearContents();

    // Check if data is undefined or empty, and if so, simply return after clearing the sheet
    if (!data || data.length === 0) {
        return; // Exit the function here; the sheet has been cleared, and there's no data to process
    }
  
    // If indexesToRemove is provided and it's an array, filter the data
    if (Array.isArray(indexesToRemove) && indexesToRemove.length > 0) {
        data = data.map(item => {
            if (typeof item === 'object') { // Ensure that we are dealing with an object
                let keys = Object.keys(item);
                return keys.reduce((obj, key, index) => {
                    if (!indexesToRemove.includes(index)) {
                        obj[key] = item[key]; // keep this field in the new object
                    }
                    return obj;
                }, {});
            }
            return item; // If not an object, return the item as is
        });
    }

    // Check if data is an array of objects or a simple array
    if (data.length > 0 && typeof data[0] === 'object') {
        const header = [Object.keys(data[0])];
        const rows = data.map(row => Object.values(row));
        const allRows = header.concat(rows);
        sheet.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
    } else {
        // Handle the case for a simple array like marketingAssets
        const rows = data.map(item => [item]);
        sheet.getRange(1, 1, rows.length, 1).setValues(rows);
    }
}



// output summary data
  function outputSummaryToSheet(ss, sheetName, summaryData) {
    // Get the sheet with the name `sheetName` from `ss`. Create it if it doesn't exist.
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    sheet.clear();
    for (let i = 0; i < summaryData.length; i++) {
      sheet.getRange(i + 1, 1, 1, summaryData[i].length).setValues([summaryData[i]]);
    }
    sheet.setFrozenRows(1);
    sheet.getRange("1:1").setFontWeight("bold");
    sheet.getRange("1:1").setWrap(true);
    sheet.setRowHeight(1, 40);
  }



// create summary on one tab
  function processSummaryData(processedCampData, processedAssetGroupData, processedDisplayData, processedDataForVideo, searchResults) {
    const header = ['Date', 'Campaign Name', 
                    'Camp Cost', 'Camp Conv', 'Camp Value',
                    'Shop Cost', 'Shop Conv', 'Shop Value',
                    'Disp Cost', 'Disp Conv', 'Disp Value',
                    'Video Cost', 'Video Conv', 'Video Value',
                    'Search Cost', 'Search Conv', 'Search Value'];

    const summaryData = {};

    // Helper function to update summaryData
    function updateSummary(row, type) {
      const date = row['date'];
      const campaignName = row['campaignName'];
      const key = `${date}_${campaignName}`;

      // Initialize if the key doesn't exist
      if (!summaryData[key]) {
        summaryData[key] = {
          'date': date,
          'campaignName': campaignName,
          'generalCost': 0, 'generalConv': 0, 'generalConvValue': 0,
          'shoppingCost': 0, 'shoppingConv': 0, 'shoppingConvValue': 0,
          'displayCost': 0, 'displayConv': 0, 'displayConvValue': 0,
          'videoCost': 0, 'videoConv': 0, 'videoConvValue': 0,
          'searchCost': 0, 'searchConv': 0, 'searchConvValue': 0
        };
      }

      // Update the metrics
      summaryData[key][`${type}Cost`] += row['cost'];
      summaryData[key][`${type}Conv`] += row['conversions'];
      summaryData[key][`${type}ConvValue`] += row['conversionsValue'];
    }

    // Process each lot of data
    processedCampData.forEach(row => updateSummary(row, 'general'));
    processedAssetGroupData.forEach(row => updateSummary(row, 'shopping'));
    processedDisplayData.forEach(row => updateSummary(row, 'display'));
    processedDataForVideo.forEach(row => updateSummary(row, 'video'));
    searchResults.forEach(row => updateSummary(row, 'search'));

    // Convert summaryData to an array format
    const summaryArray = Object.values(summaryData).map(summaryRow => {
      return [
        summaryRow.date,
        summaryRow.campaignName,
        summaryRow.generalCost, summaryRow.generalConv, summaryRow.generalConvValue,
        summaryRow.shoppingCost, summaryRow.shoppingConv, summaryRow.shoppingConvValue,
        summaryRow.displayCost, summaryRow.displayConv, summaryRow.displayConvValue,
        summaryRow.videoCost, summaryRow.videoConv, summaryRow.videoConvValue,
        summaryRow.searchCost, summaryRow.searchConv, summaryRow.searchConvValue
      ];
    });

    return [header, ...summaryArray];
  }



// Process total data for 'Totals' tab
  function processTotalData(processedCampData, processedAssetGroupData, processedDisplayData, processedDataForVideo, searchResults) {
    const header = ['Campaign Name', 
                    'Camp Cost', 'Camp Conv', 'Camp Value',
                    'Shop Cost', 'Shop Conv', 'Shop Value',
                    'Disp Cost', 'Disp Conv', 'Disp Value',
                    'Video Cost', 'Video Conv', 'Video Value',
                    'Search Cost', 'Search Conv', 'Search Value'];

    const totalData = {};

    // Helper function to update totalData
    function updateTotal(row, type) {
      const campaignName = row['campaignName'];
      const key = campaignName;

      // Initialize if the key doesn't exist
      if (!totalData[key]) {
        totalData[key] = {
          'campaignName': campaignName,
          'generalCost': 0, 'generalConv': 0, 'generalConvValue': 0,
          'shoppingCost': 0, 'shoppingConv': 0, 'shoppingConvValue': 0,
          'displayCost': 0, 'displayConv': 0, 'displayConvValue': 0,
          'videoCost': 0, 'videoConv': 0, 'videoConvValue': 0,
          'searchCost': 0, 'searchConv': 0, 'searchConvValue': 0
        };
      }

      // Update the metrics
      totalData[key][`${type}Cost`] += row['cost'];
      totalData[key][`${type}Conv`] += row['conversions'];
      totalData[key][`${type}ConvValue`] += row['conversionsValue'];
    }

    // Process each lot of data
    processedCampData.forEach(row => updateTotal(row, 'general'));
    processedAssetGroupData.forEach(row => updateTotal(row, 'shopping'));
    processedDisplayData.forEach(row => updateTotal(row, 'display'));
    processedDataForVideo.forEach(row => updateTotal(row, 'video'));
    searchResults.forEach(row => updateTotal(row, 'search'));

    // Convert totalData to an array format
    const totalArray = Object.values(totalData).map(totalRow => {
      return [
        totalRow.campaignName,
        totalRow.generalCost, totalRow.generalConv, totalRow.generalConvValue,
        totalRow.shoppingCost, totalRow.shoppingConv, totalRow.shoppingConvValue,
        totalRow.displayCost, totalRow.displayConv, totalRow.displayConvValue,
        totalRow.videoCost, totalRow.videoConv, totalRow.videoConvValue,
        totalRow.searchCost, totalRow.searchConv, totalRow.searchConvValue
      ];
    });

    return [header, ...totalArray];
  }



// get product data & iterate through the three use cases
function fetchProductData(queryString, outputType) {
  let data = [];
  let aggregatedData = {};
  const iterator = AdsApp.search(queryString);

  while (iterator.hasNext()) {
    const row = iterator.next();
    let rowData = flattenObject(row); 

    // Remove unwanted fields
    delete rowData['campaign.resourceName'];
    delete rowData['campaign.advertisingChannelType'];
    delete rowData['shoppingPerformanceView.resourceName'];

    // Determine unique key based on output type
    let uniqueKey;
    switch (outputType) {
      case 'products':
        uniqueKey = rowData['segments.productTitle'];
        break;
      case 'prod2':
        uniqueKey = rowData['segments.productTitle'] + '|' + rowData['segments.productItemId'];
        break;
      case 'id':
        uniqueKey = rowData['segments.productItemId'];
        break;
      default:
        throw new Error("Invalid output type");
    }

    // Aggregate data
    if (!aggregatedData.hasOwnProperty(uniqueKey)) {
      aggregatedData[uniqueKey] = {};
    }
      aggregatedData[uniqueKey]['Impr'] = (aggregatedData[uniqueKey]['Impr'] || 0) + (Number(rowData['metrics.impressions']) || 0);
        aggregatedData[uniqueKey]['Clicks'] = (aggregatedData[uniqueKey]['Clicks'] || 0) + (Number(rowData['metrics.clicks']) || 0);
        aggregatedData[uniqueKey]['Cost'] = (aggregatedData[uniqueKey]['Cost'] || 0) + ((Number(rowData['metrics.costMicros']) / 1e6) || 0);
        aggregatedData[uniqueKey]['Conv'] = (aggregatedData[uniqueKey]['Conv'] || 0) + (Number(rowData['metrics.conversions']) || 0);
        aggregatedData[uniqueKey]['Value'] = (aggregatedData[uniqueKey]['Value'] || 0) + (Number(rowData['metrics.conversionsValue']) || 0);
        aggregatedData[uniqueKey]['Product Title'] = rowData['segments.productTitle'];
        aggregatedData[uniqueKey]['Product ID'] = rowData['segments.productItemId'];
    } // end while

  // Post-processing for additional fields and calculations
  for (let key in aggregatedData) {
        // Calculate ROAS, CvR & CTR
        aggregatedData[key]['ROAS'] = (aggregatedData[key]['Cost'] > 0) ? aggregatedData[key]['Value'] / aggregatedData[key]['Cost'] : 0;
        aggregatedData[key]['CvR'] = (aggregatedData[key]['Clicks'] === 0) ? 0 : aggregatedData[key]['Conv'] / aggregatedData[key]['Clicks'];
        aggregatedData[key]['CTR'] = (aggregatedData[key]['Clicks'] === 0) ? 0 : aggregatedData[key]['Clicks'] / aggregatedData[key]['Impr'];

    // Assign buckets
    if (aggregatedData[key]['Cost'] === 0) {
      aggregatedData[key]['Bucket'] = 'zombie';
    } else if (aggregatedData[key]['Conv'] === 0) {
      aggregatedData[key]['Bucket'] = 'zeroconv';
    } else {
      if (aggregatedData[key]['Cost'] < tCost) {
        if (aggregatedData[key]['ROAS'] < tRoas) {
          aggregatedData[key]['Bucket'] = 'meh';
        } else {
          aggregatedData[key]['Bucket'] = 'flukes';
        }
      } else {
        if (aggregatedData[key]['ROAS'] < tRoas) {
          aggregatedData[key]['Bucket'] = 'costly';
        } else {
          aggregatedData[key]['Bucket'] = 'profitable';
        }
      }
    }
  
    let baseDataObject = {
        'Product Title': aggregatedData[key]['Product Title'],
        'Product ID': aggregatedData[key]['Product ID'],
        'Impr': aggregatedData[key]['Impr'],
        'Clicks': aggregatedData[key]['Clicks'],
        'Cost': aggregatedData[key]['Cost'],
        'Conv': aggregatedData[key]['Conv'],
        'Value': aggregatedData[key]['Value'],
        'CTR': aggregatedData[key]['CTR'],
        'ROAS': aggregatedData[key]['ROAS'],
        'CvR': aggregatedData[key]['CvR'],
        'Bucket': aggregatedData[key]['Bucket']
    };

    if (outputType === 'products') {
        delete baseDataObject['Product ID']; // Remove the Product ID for 'products' output type
    } else if (outputType === 'id') {
        baseDataObject = {
            'Product ID': aggregatedData[key]['Product ID'],
            'Bucket': aggregatedData[key]['Bucket']
        };
    }

    data.push(baseDataObject);
    
    } // end for loop

    return data;
}



// update variables from the settings tab
function updateVariablesFromSheet(ss, sheetName) {
  // Define the names of your named ranges
  const numberOfDaysRangeName = 'numberOfDays';
  const tCostRangeName = 'tCost';
  const tRoasRangeName = 'tRoas';
  const productDaysRangeName = 'productDays';
  const brandTermRangeName = 'brandTerm'; 

  try {
    // Access the specific sheet within the spreadsheet
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return;
    }

    // Fetch the ranges by name from the specific sheet
    const numberOfDaysRange = sheet.getRange(numberOfDaysRangeName);
    const tCostRange = sheet.getRange(tCostRangeName);
    const tRoasRange = sheet.getRange(tRoasRangeName);
    const productDaysRange = sheet.getRange(productDaysRangeName);
    const brandTermRange = sheet.getRange(brandTermRangeName); 

    // Check if the range is found and the value is not empty
    // Then update the global variables 
    if (numberOfDaysRange && numberOfDaysRange.getValue() !== "") {
      const value = numberOfDaysRange.getValue();
      if (!isNaN(value)) {
        numberOfDays = value;
      }
    }

    if (tCostRange && tCostRange.getValue() !== "") {
      const value = tCostRange.getValue();
      if (!isNaN(value)) {
        tCost = value;
      }
    }

    if (tRoasRange && tRoasRange.getValue() !== "") {
      const value = tRoasRange.getValue();
      if (!isNaN(value)) {
        tRoas = value;
      }
    }

    if (productDaysRange && productDaysRange.getValue() !== "") {
      const value = productDaysRange.getValue();
      if (!isNaN(value)) {
        productDays = value;
      }
    }

    // Check if 'brandTermRange' exists and its value is non-empty
    const value = brandTermRange && brandTermRange.getValue();
    if (value && typeof value === 'string' && value.trim() !== "") {
        brandTerm = value; // update the global variable
    } else if (!brandTerm || (typeof brandTerm === 'string' && brandTerm.trim() === "")) {
        brandTerm = "brand is not entered in script or client sheet"; // default value
        console.log(`${brandTerm}`);
    }


    
  } catch (e) {
    // If there's an error, we simply don't update the variables
    // Log the error for debugging purposes
    console.error("Error in updateVariablesFromSheet:", e);
  }
}



// extract search terms
  function extractTerms(ss, mainDateRange, tol) {
      // Extract Campaign IDs with Status not 'REMOVED' and impressions > 0
      let campaignIdsQuery = AdsApp.report(`
      SELECT campaign.id, campaign.name, metrics.clicks, 
          metrics.impressions, metrics.conversions, metrics.conversions_value
          FROM campaign 
          WHERE campaign.status != 'REMOVED' 
          AND campaign.advertising_channel_type = "PERFORMANCE_MAX" 
          AND metrics.impressions > 0 AND ${mainDateRange} 
          ORDER BY metrics.conversions DESC `
      );

      let rows = campaignIdsQuery.rows();   
      let allSearchTerms = [['Campaign Name', 'Campaign ID', 'Category Label', 'Clicks', 'Impr', 'Conv', 'Value', 'Bucket', 'Distance']];

      while (rows.hasNext()) {
          let row = rows.next();

          // Fetch search terms and related metrics for the campaign ordered by conversions descending
          let query = AdsApp.report(` 
          SELECT campaign_search_term_insight.category_label, metrics.clicks, 
              metrics.impressions, metrics.conversions, metrics.conversions_value  
              FROM campaign_search_term_insight 
              WHERE ${mainDateRange}
              AND campaign_search_term_insight.campaign_id = ${row['campaign.id']} 
              ORDER BY metrics.impressions DESC `
          );

          let searchTermRows = query.rows();
          while (searchTermRows.hasNext()) {
              let searchTermRow = searchTermRows.next();
              let term = searchTermRow['campaign_search_term_insight.category_label'];
              let { bucket, distance } = determineBucketAndDistance(term, tol); // Destructure to get bucket and distance
              allSearchTerms.push([row['campaign.name'], row['campaign.id'], 
                                   term,
                                   searchTermRow['metrics.clicks'], 
                                   searchTermRow['metrics.impressions'], 
                                   searchTermRow['metrics.conversions'], 
                                   searchTermRow['metrics.conversions_value'],
                                   bucket, // Include the bucket in your data
                                   distance]); // Include the distance in your data
          }

      }

      // Write all search terms to the 'terms' sheet 
      let termsSheet = ss.getSheetByName('terms') ? ss.getSheetByName('terms').clear() : ss.insertSheet('terms');
      termsSheet.getRange(1, 1, allSearchTerms.length, allSearchTerms[0].length).setValues(allSearchTerms);


      // Aggregate terms and write to the 'totalTerms' sheet
      aggregateTerms(allSearchTerms, ss);
  }



// aggregate search terms
function aggregateTerms(searchTerms, ss) {
    let aggregated = {}; // { term: { clicks: 0, impressions: 0, conversions: 0, conversionValue: 0, bucket: '', distance: 0 }, ... }

    for (let i = 1; i < searchTerms.length; i++) { // Start from 1 to skip headers
        let term = searchTerms[i][2] || 'blank'; // Use 'blank' for empty search terms

        if (!aggregated[term]) {
            aggregated[term] = { 
                clicks: 0, 
                impressions: 0, 
                conversions: 0, 
                conversionValue: 0, 
                bucket: searchTerms[i][7],  // Assuming bucket is in the 8th position of your array
                distance: searchTerms[i][8]  // Assuming distance is in the 9th position of your array
            };
        }

        aggregated[term].clicks += Number(searchTerms[i][3]);
        aggregated[term].impressions += Number(searchTerms[i][4]);
        aggregated[term].conversions += Number(searchTerms[i][5]);
        aggregated[term].conversionValue += Number(searchTerms[i][6]);
        // Assuming that the bucket and distance are the same for all instances of a term,
        // we don't aggregate them but just take the value from the first instance.
    }

    let aggregatedArray = [['Category Label', 'Clicks', 'Impressions', 'Conversions', 'Conversion Value', 'Bucket', 'Distance']]; // Header row
    for (let term in aggregated) {
        aggregatedArray.push([
            term, 
            aggregated[term].clicks, 
            aggregated[term].impressions, 
            aggregated[term].conversions, 
            aggregated[term].conversionValue, 
            aggregated[term].bucket,  // Adding bucket to output
            aggregated[term].distance  // Adding distance to output
        ]);
    }

    let header = aggregatedArray.shift(); // Remove the header before sorting
    // Sort by impressions descending
    aggregatedArray.sort((a, b) => b[2] - a[2]);
    aggregatedArray.unshift(header); // Prepend the header back to the top

    // Write aggregated data to the 'totalTerms' sheet
    let totalTermsSheet = ss.getSheetByName('totalTerms') ? ss.getSheetByName('totalTerms').clear() : ss.insertSheet('totalTerms');
    totalTermsSheet.getRange(1, 1, aggregatedArray.length, aggregatedArray[0].length).setValues(aggregatedArray);
}



// find single word nGrams
function extractNGrams(ss, mainDateRange) {
    // Extract Campaign IDs with Status not 'REMOVED' and impressions > 0
    let campaignIdsQuery = AdsApp.report(`SELECT campaign.id, campaign.name, metrics.clicks, metrics.impressions, metrics.conversions, metrics.conversions_value
        FROM campaign WHERE campaign.status != 'REMOVED' AND metrics.impressions > 0 AND ${mainDateRange} ORDER BY metrics.conversions DESC `
    );

    let rows = campaignIdsQuery.rows();
    let campaignNGrams = {}; // To store nGram data per campaign

    while (rows.hasNext()) {
        let row = rows.next();
        let campaignId = row['campaign.id'];
        let campaignName = row['campaign.name'];

        // Fetch search terms and related metrics for the campaign ordered by conversions descending
        let query = AdsApp.report(` SELECT campaign_search_term_insight.category_label, metrics.clicks, metrics.impressions, 
            metrics.conversions, metrics.conversions_value  FROM campaign_search_term_insight WHERE ${mainDateRange}
            AND campaign_search_term_insight.campaign_id = ${campaignId} AND metrics.impressions > 0 ORDER BY metrics.impressions DESC `
        );

        let searchTermRows = query.rows();
        while (searchTermRows.hasNext()) {
            let searchTermRow = searchTermRows.next();
            let terms = searchTermRow['campaign_search_term_insight.category_label'].split(' '); // Split the category label into individual words

            terms.forEach((term) => {
                let key = campaignId + '_' + term; // Unique key for each nGram within a campaign

                if (!campaignNGrams[key]) {
                    campaignNGrams[key] = {
                        campaignName: campaignName,
                        campaignId: campaignId,
                        nGram: term,
                        clicks: 0,
                        impressions: 0,
                        conversions: 0,
                        conversionValue: 0
                    };
                }

                // Aggregate the metrics for the nGram within the campaign
                campaignNGrams[key].clicks += Number(searchTermRow['metrics.clicks']);
                campaignNGrams[key].impressions += Number(searchTermRow['metrics.impressions']);
                campaignNGrams[key].conversions += Number(searchTermRow['metrics.conversions']);
                campaignNGrams[key].conversionValue += Number(searchTermRow['metrics.conversions_value']);
            });
        }
    }

    // Convert the aggregated object to an array format suitable for the spreadsheet
    let allNGrams = [['Campaign Name', 'Campaign ID', 'nGram', 'Clicks', 'Impressions', 'Conversions', 'Conversion Value']];
    for (let key in campaignNGrams) {
        let item = campaignNGrams[key];
        allNGrams.push([item.campaignName, item.campaignId, item.nGram, item.clicks, item.impressions, item.conversions, item.conversionValue]);
    }

    // Write all nGrams to the 't2' sheet
    let nGramsSheet = ss.getSheetByName('t2') ? ss.getSheetByName('t2').clear() : ss.insertSheet('t2');
    nGramsSheet.getRange(1, 1, allNGrams.length, allNGrams[0].length).setValues(allNGrams);

    // Continue with aggregation for 'total2' as before
    aggregateNGrams(allNGrams, ss);
}



// aggregate the nGrams
function aggregateNGrams(nGrams, ss) {
    // Similar structure to aggregateTerms, but we're aggregating nGrams
    let aggregated = {};

    for (let i = 1; i < nGrams.length; i++) {
        let nGram = nGrams[i][2] || 'blank';

        if (!aggregated[nGram]) {
            aggregated[nGram] = { clicks: 0, impressions: 0, conversions: 0, conversionValue: 0 };
        }

        aggregated[nGram].clicks += Number(nGrams[i][3]);
        aggregated[nGram].impressions += Number(nGrams[i][4]);
        aggregated[nGram].conversions += Number(nGrams[i][5]);
        aggregated[nGram].conversionValue += Number(nGrams[i][6]);
    }

    let aggregatedArray = [['nGram', 'Clicks', 'Impressions', 'Conversions', 'Conversion Value']];
    for (let nGram in aggregated) {
        aggregatedArray.push([nGram, aggregated[nGram].clicks, aggregated[nGram].impressions, aggregated[nGram].conversions, aggregated[nGram].conversionValue]);
    }

    let header = aggregatedArray.shift();
    aggregatedArray.sort((a, b) => b[2] - a[2]); // Sort by impressions
    aggregatedArray.unshift(header);

    // Write aggregated data to the 'total2' sheet
    let totalNGramsSheet = ss.getSheetByName('total2') ? ss.getSheetByName('total2').clear() : ss.insertSheet('total2');
    totalNGramsSheet.getRange(1, 1, aggregatedArray.length, aggregatedArray[0].length).setValues(aggregatedArray);
}




// Helper function to determine the bucket for a term and calculate Levenshtein distance
function determineBucketAndDistance(term, tolerance) {
  // Check if 'brandTerm' is defined and non-empty
  if (typeof brandTerm === 'undefined' || brandTerm === null || brandTerm.trim() === '') {
    console.error("Brand term is not defined or empty.");
    // Handle the error appropriately, e.g., by returning a default value or stopping execution
    return { bucket: 'unknown', distance: '' }; // or throw an error
  }

  const lowerCaseTerm = term.toLowerCase();
  const lowerCaseBrandTerm = brandTerm.toLowerCase();

  let distance = lowerCaseTerm === 'blank' || lowerCaseTerm === '' ? '' : levenshtein(lowerCaseTerm, lowerCaseBrandTerm);

  if (lowerCaseTerm === 'blank' || lowerCaseTerm === '') {
    return { bucket: 'blank', distance: '' };
  } else if (lowerCaseTerm === lowerCaseBrandTerm) {
    return { bucket: 'brand', distance };
  } else {
    // Check if the term contains the brand term or a close variant
    if (lowerCaseTerm.includes(lowerCaseBrandTerm)) {
      return { bucket: 'close-brand', distance };
    } else {      
      // Check for a close match with a tolerance defined by levenshteinTolerance
      if (distance <= tolerance) {
        return { bucket: 'close-brand', distance };
      } else {
        return { bucket: 'non-brand', distance };
      }
    }
  }
}



// Function to calculate Levenshtein distance between two strings
function levenshtein(a, b) {
  var tmp;
  if (a.length === 0) { return b.length; }
  if (b.length === 0) { return a.length; }
  if (a.length > b.length) { tmp = a; a = b; b = tmp; }

  var i, j, res, alen = a.length, blen = b.length, row = Array(alen);
  for (i = 0; i <= alen; i++) { row[i] = i; }

  for (i = 1; i <= blen; i++) {
    res = i;
    for (j = 1; j <= alen; j++) {
      tmp = row[j - 1];
      row[j - 1] = res;
      res = b[i - 1] === a[j - 1] ? tmp : Math.min(tmp + 1, Math.min(res + 1, row[j] + 1));
    }
  }
  return res;
}
