function main() {
  try {
    const spreadsheetId = 'YOUR_SPREADSHEET_ID';
    Logger.log('Attempting to open spreadsheet with ID: ' + spreadsheetId);
    let ss = SpreadsheetApp.openById(spreadsheetId);
    Logger.log('Spreadsheet opened successfully.');

    let zombieDays = 366;
    let prodDays = 181;

    // Define query elements. wrap with spaces for safety
    let impr = ' metrics.impressions ';
    let clicks = ' metrics.clicks ';
    let cost = ' metrics.cost_micros ';
    let conv = ' metrics.conversions '; 
    let value = ' metrics.conversions_value '; 
    let allConv = ' metrics.all_conversions '; 
    let allValue = ' metrics.all_conversions_value '; 
    let views = ' metrics.video_views ';
    let cpv = ' metrics.average_cpv ';
    let segDate = ' segments.date ';  
    let prodTitle = ' segments.product_title ';
    let prodID = ' segments.product_item_id ';
    let prodC0 = ' segments.product_custom_attribute0 ';
    let prodC1 = ' segments.product_custom_attribute1 ';
    let prodC2 = ' segments.product_custom_attribute2 ';
    let prodC3 = ' segments.product_custom_attribute3 ';
    let prodC4 = ' segments.product_custom_attribute4 '; 
    let campName = ' campaign.name ';
    let chType = ' campaign.advertising_channel_type ';
    let adgName = ' ad_group.name ';
    let adStatus = ' ad_group_ad.status ';
    let adPerf = ' ad_group_ad_asset_view.performance_label ';
    let adType = ' ad_group_ad_asset_view.field_type ';
    let aIdAsset = ' asset.resource_name ';  
    let aId = ' asset.id ';
    let assetType = ' asset.type ';
    let aFinalUrl = ' asset.final_urls ';
    let assetName = ' asset.name ';
    let assetText = ' asset.text_asset.text ';
    let assetSource = ' asset.source ' ; 
    let adUrl = ' asset.image_asset.full_size.url ';
    let ytTitle = ' asset.youtube_video_asset.youtube_video_title ';
    let ytId = ' asset.youtube_video_asset.youtube_video_id ';
    let agId = ' asset_group.id ';    
    let assetFtype = ' asset_group_asset.field_type ';
    let adPmaxPerf = ' asset_group_asset.performance_label ';  
    let agStrength = ' asset_group.ad_strength ';
    let agStatus = ' asset_group.status ';
    let asgName = ' asset_group.name ';
    let lgType = ' asset_group_listing_group_filter.type ';  
    let aIdCamp = ' segments.asset_interaction_target.asset ';
    let interAsset = ' segments.asset_interaction_target.interaction_on_this_asset ';
    let pMaxOnly = ' AND campaign.advertising_channel_type = "PERFORMANCE_MAX" '; 
    let searchOnly = ' AND campaign.advertising_channel_type = "SEARCH" ';   
    let agFilter = ' AND asset_group_listing_group_filter.type != "SUBDIVISION" ';   
    let adgEnabled = ' AND ad_group.status = "ENABLED" AND campaign.status = "ENABLED" AND ad_group_ad.status = "ENABLED" ';
    let asgEnabled = ' asset_group.status = "ENABLED" AND campaign.status = "ENABLED" ';           
    let notInter = ' AND segments.asset_interaction_target.interaction_on_this_asset != "TRUE" ';
    let inter = ' AND segments.asset_interaction_target.interaction_on_this_asset = "TRUE" ';
    let date07 = ' segments.date DURING LAST_7_DAYS ';  
    let date30 = ' segments.date DURING LAST_30_DAYS ';  
    let order = ' ORDER BY campaign.name '; 
    let orderImpr = ' ORDER BY metrics.impressions DESC '; 
  
    // Date stuff
    let MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    let now = new Date();
    let from = new Date(now.getTime() - zombieDays * MILLIS_PER_DAY);        
    let prod180 = new Date(now.getTime() - prodDays * MILLIS_PER_DAY);       
    let to = new Date(now.getTime() - 1 * MILLIS_PER_DAY);                   
    let timeZone = AdsApp.currentAccount().getTimeZone(); 
    let zombieRange = ' segments.date BETWEEN "' + Utilities.formatDate(from, timeZone, 'yyyy-MM-dd') + '" AND "' + Utilities.formatDate(to, timeZone, 'yyyy-MM-dd') + '"' 
    let prodDate = ' segments.date BETWEEN "' + Utilities.formatDate(prod180, timeZone, 'yyyy-MM-dd') + '" AND "' + Utilities.formatDate(to, timeZone, 'yyyy-MM-dd') + '"'
  
    // Build queries                     
    let cd = [segDate, campName, cost, conv, value, views, cpv, impr, clicks, chType]; 
    let campQuery = 'SELECT ' + cd.join(',') + 
      ' FROM campaign ' +
      ' WHERE ' + date30 + pMaxOnly + order; 
  
    let dv = [segDate, campName, aIdCamp, cost, conv, value, views, cpv, impr, chType, interAsset]; 
    let dvQuery = 'SELECT ' + dv.join(',') + 
      ' FROM campaign ' +
      ' WHERE ' + date30 + pMaxOnly + notInter + order; 
  
    let p = [campName, prodTitle, cost, conv, value, impr, chType, prodID, prodC0, prodC1, prodC2, prodC3, prodC4]; 
    let pQuery = 'SELECT ' + p.join(',')  + 
      ' FROM shopping_performance_view  ' + 
      ' WHERE ' + date30 + pMaxOnly + order; 
    let p180Query = 'SELECT ' + p.join(',')  + 
      ' FROM shopping_performance_view  ' + 
      ' WHERE ' + prodDate + pMaxOnly + order;   

    let ag = [segDate, campName, asgName, agStrength, agStatus, lgType, impr, clicks, cost, conv, value]; 
    let agQuery = 'SELECT ' + ag.join(',')  + 
      ' FROM asset_group_product_group_view ' +
      ' WHERE ' + date30 + agFilter;

    let assets = [aId, aFinalUrl, assetSource, assetType, ytTitle, ytId, assetText, aIdAsset, assetName]; 
    let assetQuery = 'SELECT ' + assets.join(',')  + 
      ' FROM asset';

    let ads = [campName, asgName, agId, aIdAsset, assetFtype, adPmaxPerf, agStrength, agStatus, assetSource]; 
    let adsQuery = 'SELECT ' + ads.join(',') +
      ' FROM asset_group_asset';

    let zombies = [prodID, clicks, impr, prodTitle]; 
    let zQuery = 'SELECT ' + zombies.join(',') +
      ' FROM shopping_performance_view ' +
      ' WHERE metrics.clicks < 1 AND ' + zombieRange + orderImpr;  

    // Call report function to pull data & push it to the named tabs in the sheet
    runReport(campQuery, ss.getSheetByName('r_camp'));  
    runReport(dvQuery, ss.getSheetByName('r_dv'));     
    runReport(pQuery, ss.getSheetByName('r_prod_t')); 
    runReport(p180Query, ss.getSheetByName('r_prod_t_180'));   
    runReport(agQuery, ss.getSheetByName('r_ag'));   
    runReport(assetQuery, ss.getSheetByName('r_allads'));
    runReport(adsQuery, ss.getSheetByName('r_ads'));   
    runReport(zQuery, ss.getSheetByName('zombies')); 
  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}

// Query & export report data to named sheet
function runReport(q, sh) {
  try {
    Logger.log('Running query: ' + q);
    const report = AdsApp.report(q);
    Logger.log('Exporting to sheet: ' + sh.getName());
    report.exportToSheet(sh);  
  } catch (e) {
    Logger.log('Error in runReport: ' + e.message + ' Query: ' + q);
  }
}
