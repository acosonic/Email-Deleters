var email_search = 'category:updates label:unread';
var MAX_RETRIES = 3;

function batchDeleteEmail() {
  _run(0);
}

function _run(retryCount) {
  try {
    processEmail(email_search, 'moveThreadsToTrash');
  } catch (error) {
    Logger.log('Error: ' + error);
    if (retryCount < MAX_RETRIES) {
      Logger.log('Re-trying (attempt ' + (retryCount + 1) + ' of ' + MAX_RETRIES + ')...');
      _run(retryCount + 1);
    } else {
      Logger.log('Max retries reached, giving up.');
    }
  }
}

function processEmail(search, batchAction) {
  var removed = 0;
  var batchSize = 100; // Process up to 100 threads at once
  var searchSize = 500; // Maximum search result size is 500

  var threads = GmailApp.search(search, 0, searchSize);
  Logger.log('Found ' + threads.length + ' threads matching search: "' + search + '".');

  while (threads.length > 0) {
    for (var j = 0; j < threads.length; j += batchSize) {
      var batch = threads.slice(j, j + batchSize);
      Logger.log('Removing threads ' + j + ' to ' + (j + batch.length) + '...');
      GmailApp[batchAction](batch);
      removed += batch.length;

      var rando = Math.floor(Math.random() * 3000);
      Logger.log('Sleeping for ' + rando + ' milliseconds.');
      Utilities.sleep(rando);
    }

    if (removed > 0 && removed % 2000 === 0) {
      var longSleep = Math.floor(Math.random() * 20000) + 20000;
      Logger.log('Long sleeping for ' + longSleep + ' milliseconds.');
      Utilities.sleep(longSleep);
    }

    Logger.log('Total removed so far: ' + removed + '.');
    threads = GmailApp.search(search, 0, searchSize);
    Logger.log('Found ' + threads.length + ' more threads.');
  }

  Logger.log('Done. Total removed: ' + removed + '.');
}
