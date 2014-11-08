var moment = require('moment');

var jsDate = function(days,epoch){
  var startEpoch;
  if (epoch === 1900){
    startEpoch = moment('1900-1-1','YYYY-M-D');
  } else {
    startEpoch = moment('1904-1-1','YYYY-M-D');
  }
  var newDate = startEpoch.add(days,'days');
  return newDate.toDate().toString();
};

module.exports = jsDate;
