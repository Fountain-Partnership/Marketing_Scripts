/*************************************************
 * Bing Alerts - v1.0
 * By: Fountain (@FountainTeam)
 * Contributors:
 *  Ephraim
 *
 * Usage:
 *  let settingsSheetPointer = 'URL_FOR_SPREADSHEET';
 *  let spreadsheetUrl = 'URL_FOR_CSV';
 *  let slack_channel_name = '#general';
 *  let slack_webhook_url = 'https://hooks.slack.com/services/~~/~~/~~/';
 *
 *  main: \{...}\
 *
 * Description:
 *  A bing ads script that uses a google spreadsheet
 *  and slack to send performance reports.
 *  Instructions here: bit.ly/fountainbingalerts/
 *
 *************************************************/

function main()
{
    let settingsSheetPointer = 'URL_FOR_SPREADSHEET';
    let spreadsheetUrl = 'URL_FOR_CSV';
    let slack_channel_name = '#general';
    let slack_webhook_url = 'https://hooks.slack.com/services/YOUR/WEBHOOK/ID-GOES-HERE';

    let reportingConfig = parseCsv(spreadsheetUrl, true);

    main:
    {

      try
      {
          var accounts = AccountsApp.accounts().get();
          var hasAccounts = true;
      } catch(e)
      {
          var hasAccounts = false;
      }

      if (hasAccounts)
      {
          var reportObject = new Report(reportingConfig, true);

      }
      else
      {
          var reportObject = new Report(reportingConfig, false);

      }

      if (!reportObject.isAgencyShellReport)
      {
          var report = reportObject.reportValues.reportRows.join("\n");

          report = report + "\nTo view or change your settings, see the spreadsheet: "+ settingsUrl;


          sendSlackMessage(report, slack_channel_name, slack_webhook_url);

      }
      else
      {
          var reportObject = reportObject.reportValues;
          var report = ["*Agency Shell Alerts:*"];
          for (var i = 0 ; i < reportObject.length; i++)
          {
              var accountReport = reportObject[i];
              var reportRows = accountReport.reportRows;

              if (reportRows.length > 0)
              {
                var accountName = accountReport.accountName;
                reportRows.unshift(`- ${accountName}:`);
                reportRows = reportRows.join("\n\t\t");
                report.push(reportRows);
              }
          }
          report = report.join("\n");
          report = report + "\nTo view or change your settings, see the spreadsheet: "+ settingsUrl;
          sendSlackMessage(report, slack_channel_name, slack_webhook_url);

      }
    }
}


function parseCsv(url, row1HasHeaders)
{
    let retVal = {};
    let fetch = UrlFetchApp.fetch(url);
    let responseCode = fetch.getResponseCode();

    // validation
    if (responseCode == 200)
    {
      var csvContent = fetch.getContentText();

      if (csvContent.indexOf("<!DOCTYPE html>") > -1)
      {
        throw "The script could not parse your csv: the URL provided returns "+
        "a html page. This could be because you have not published the "+
        "spreadsheet publically. Please check your settings and run the "+
        "script again.";
      }
    }
    else
    {
        throw `bad url, response code = ${responseCode}`;
    }
    csvContent = csvContent.split("\n");

    var rows = new csviterator(csvContent);

    if (row1HasHeaders)
    {
        // skip the first row
        let headers = rows.next();
    }
    // creating a return object that will be used later in the script to
    // write reports.
    while (rows.hasNext())
    {
        let row = rows.next();
        let [accountName, entityType, metric, dateRange, operator, value] = row;

        if (accountName == "") accountName = 'Current Account';

        if (row.indexOf("") < 1)
        {
            if (!(accountName in retVal))
            {
                retVal[accountName] = {};
            }
            retVal[accountName][metric] =
            {
                entityType:entityType,
                dateRange:dateRange,
                operator:operator,
                value:value
            }
        }
    }

    if (Object.keys(retVal).length > 0)
    {
        return retVal;
    } else
    {
        throw `The correct settings could not be be found in your spreadsheet. Please check your url: ${url}`;
    }

    function csviterator(csvContent)
    {
        let rows = csvContent;
        this.hasNext = function()
        {
            return rows.length > 0;
        }
        this.next = function()
        {
            if (rows.length > 0)
            {
                let retVal = rows[0].split(",");
                rows.splice(0,1);
                return retVal;
            } else
            {
                throw "The iterator has reached its end.";
            }
        }
    }
}


function Report (config, agencyShell)
{
  // the function required for agency shell reports is different from that
  // for single accounts.accountName
  if (agencyShell)
  {
    this.isAgencyShellReport = true;
    this.listOfReports = [];
    for (accountName in config)
    {
      this.listOfReports.push(checkAccountAgainstMetricsAgencyShell
        (accountName, config));
    }
    this.reportValues = this.listOfReports;

    delete this.listOfReports;
  }
  else
  {
    this.reportValues = checkAccountAgainstMetrics(config);
  }
}

function checkAccountAgainstMetricsAgencyShell(accountName, config)
{
  let retVal = {reportRows: []};
  if (accountName in config && accountName !== "Current Account")
  {
    let accountObject = config[accountName];

    var accountSearch = AccountsApp.accounts()
      .withCondition("Impressions > 1")
      .withCondition(`Name = '${accountName}'`)
      .forDateRange("YESTERDAY")
      .get();


    if (accountSearch.hasNext())
    {
      while (accountSearch.hasNext())
      {
        AccountsApp.select(accountSearch.next());

        for (metric in accountObject)
        {
          let settings = accountObject[metric];

          switch (settings.entityType)
          {
            case 'Campaign' :
                  var entities = AdsApp
                    .campaigns()
                    .withCondition("Impressions > 1")
                    .withCondition("DeliveryStatus = ELIGIBLE")
                    .forDateRange(settings.dateRange)
                    .get();
                  break;
            case 'Ad Group' :
                  var entities = AdsApp
                    .adGroups()
                    .withCondition("Impressions > 1")
                    .withCondition("CampaignStatus = ENABLED")
                    .forDateRange(settings.dateRange)
                    .get();
                  break;
            case 'Keyword':
                  var entities = AdsApp
                    .keywords()
                    .withCondition("CampaignStatus != ENABLED")
                    .forDateRange(settings.dateRange)
                    .get();
                  break;
            default :
                  var entities = AccountsApp
                    .accounts()
                    .withCondition(`Name = ${accountName}`)
                    .forDateRange(settings.dateRange)
                    .get();
                  //
                  settings.entityType = 'Account';

          }

          while (entities.hasNext())
          {
            let entity = entities.next();
            let entityName = entity.getName();
            let entityType = settings.entityType;
            let stats = entity.getStats();

            switch (metric)
            {
                case 'Average Cpc' :
                    var stat = stats.getAverageCpc();
                    break;
                case 'Average Cpm' :
                    var stat = stats.getAverageCpm();
                    break;
                case 'Average Position' :
                    var stat = stats.getAveragePosition();
                    break;
                case 'Click Conversion Rate' :
                    var stat = stats.getClickConversionRate();
                    break;
                case 'Clicks' :
                    var stat = stats.getClickConversionRate();
                    break;
                case 'Converted Clicks' :
                    var stat = stats.getConvertedClicks();
                    break;
                case 'Cost' :
                    var stat = stats.getCost();
                    break;
                case 'Ctr' :
                    var stat = stats.getCtr();
                    break;
                case 'Impressions' :
                    var stat = stats.getImpressions();
                    break;
                default :
                    throw `the stat chosen for one of your rows is invalid. The `+
                    `metric '${metric}' cannot be queried in Bing Ads Script`;
            }

            var evaluationStatement =
            [
                'if (',
                'stat ',
                settings.operator,
                settings.value,
                ') ',
                'retVal.reportRows.push(',
                '"The metric \''+metric+'\' has reached a value of '+stat+', \\n\\t\\t'+
                'outside of your accepted range for the '+entityType+' \\n\\t\\t'+
                '\''+entityName+'\'"',
                ');'
            ];
            retVal.accountName = accountName;
            eval(evaluationStatement.join(""));
          }
        };
      };
      return retVal;
    } else
    {
      Logger.log(`The account '${accountName}' could not be found in your `+
      `Agency Shell.
      Please amend the name in the settings spreadsheet. The account will be `+
      `skipped.`);
      return retVal;
    }

  }
  else
  {
      throw "Please specify which accounts you would like the script to run on "+
      "if you are using it under an Agency Shell.";
      return retVal;
  }
}

function checkAccountAgainstMetrics(config)
{
  let retVal = {reportRows: []};
  let currentAccountName = AdsApp.currentAccount().getName();

  if (config["Current Account"] && config[currentAccountName])
  {
    var account = Object.assign(config[currentAccountName],config["Current Account"]);
  }
  else if (config["Current Account"])
  {
    var account = config["Current Account"];
  }
  else if (config[currentAccountName])
  {
    var account = config[currentAccountName];
  }
  else
  {
    throw "the config contains no data, please check that your spreadsheet is "+
    "populated."
  }


  for (metric in account)
  {
      let settings = account[metric];
      switch (settings.entityType)
      {
          case 'Campaign' :
              var entities = AdsApp
              .campaigns()
              .withCondition("Impressions > 1")
              .withCondition("DeliveryStatus = ELIGIBLE")
              .forDateRange(settings.dateRange)
              .get();
              break;
          case 'Ad Group' :
              var entities = AdsApp
              .adGroups()
              .withCondition("Impressions > 1")
              .withCondition("Status = ENABLED")
              .forDateRange(settings.dateRange)
              .get();
              break;
          default :
              var entities = AdsApp
              .keywords()
              .withCondition("Impressions > 1")
              .withCondition("CampaignStatus != ENABLED")
              .forDateRange(settings.dateRange)
              .get();
      }

      while (entities.hasNext())
      {
          let entity = entities.next()
          let entityName = entity.getName();
          let entityType = settings.entityType;
          let stats = entity.getStats();

          switch (metric)
          {
              case 'Average Cpc' :
                  var stat = stats.getAverageCpc();
                  break;
              case 'Average Cpm' :
                  var stat = stats.getAverageCpm();
                  break;
              case 'Average Position' :
                  var stat = stats.getAveragePosition();
                  break;
              case 'Click Conversion Rate' :
                  var stat = stats.getClickConversionRate();
                  break;
              case 'Clicks' :
                  var stat = stats.getClickConversionRate();
                  break;
              case 'Converted Clicks' :
                  var stat = stats.getConvertedClicks();
                  break;
              case 'Cost' :
                  var stat = stats.getCost();
                  break;
              case 'Ctr' :
                  var stat = stats.getCtr();
                  break;
              case 'Impressions' :
                  var stat = stats.getImpressions();
                  break;
              default :
                  throw `the stat chosen for one of your rows is invalid. `+
                  `The metric '${metric}' cannot be queried in Bing Ads Script`;
          }

          var evaluationStatement =
          [
              'if (',
              'stat ',
              settings.operator,
              settings.value,
              ') ',
              'retVal.reportRows.push(',
              '"The metric \''+metric+'\' has reached a value of '+stat+', \\n\\t\\t'+
              'outside of your accepted range for the '+entityType+' \\n\\t\\t'+
              '\''+entityName+'\'"',
              ');'
          ];
          retVal.accountName = AdsApp.currentAccount().getName();
          eval(evaluationStatement.join(""));
      }

      if(retVal.reportRows.length > 0)
      {
        retVal.reportRows[0] ="\n*"+retVal.accountName+" Alerts:*\n"+retVal.reportRows[0];
      }
  }
  return retVal;
}

function sendSlackMessage(text, opt_channel, SLACK_URL)
{

  let slackMessage = {
    text: text,
    icon_url:
        'https://www.gstatic.com/images/icons/material/product/1x/adwords_64dp.png',
    username: 'Bing Ads Script',
    link_names: 1,
    channel: opt_channel || '#general'
  };

  let options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(slackMessage)
  };
  UrlFetchApp.fetch(SLACK_URL, options);
}
