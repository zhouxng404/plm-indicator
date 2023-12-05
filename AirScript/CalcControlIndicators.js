const origSheet = Application.Sheets("问题清单");
// console.log(origSheet.Field.GetFields());
const targetSheet = Application.Sheets("管控指标");
// console.log(targetSheet.Field.GetFields());

/**
* 查询sheet所有满足条件的数据
* @param {{ Record: { GetRecords: (arg0: { PageSize: number; Offset: any; Filter: any; }) => any; }; Name: string; }} sheet
* @param {{ mode?: string; criteria?: { field: string; op: string; values: string[]; }[] | { field: string; op: string; values: string[]; }[]; }} filter
*/
function fetchAll(sheet, filter) {
  let all = [];
  let offset = null;

  while (all.length === 0 || offset) {
    let records = sheet.Record.GetRecords({
      PageSize: 1000,
      Offset: offset,
      Filter: filter
    })
    offset = records.offset;
    let rows = records.records;
    if (rows === null || rows === undefined || rows.length === 0) {
      break;
    }
    all = all.concat(records.records);
  }
  console.log('工作簿: ' + sheet.Name + ' 数据行数: ' + all.length)
  return all;
}

/**
* 排序数据集
* @param {any[]} records
* @param {string} sortField
* @param {boolean} isAsc
*/
function sortRecordByField(records, sortField, isAsc) {
  records.sort(function (a, b) {
    let o1 = a.fields[sortField];
    let o2 = b.fields[sortField];
    if (isAsc) {
      return o1.localeCompare(o2);
    }
    else {
      return o2.localeCompare(o1);
    }
  });
}

/**
* 得到两个日期间的月份
* @param {string | number | Date} startDate
* @param {string | number | Date} endDate
*/
function getMonthsBetweenDates(startDate, endDate) {
  let start = new Date(startDate);
  start.setDate(1);
  let end = new Date(endDate);
  end.setDate(1);
  let months = [];

  while (start <= end) {
    let year = start.getFullYear().toString();
    let month = start.getMonth() + 1;

    let fullMonth = `${String(month).padStart(2, '0')}`;
    let dateStr = `${year}/${fullMonth}`;
    let dateDesc = `${year}年${fullMonth}月`;

    let monthObj = {
      dateStr: dateStr,
      dateDesc: dateDesc
    };

    months.push(monthObj);

    start.setMonth(start.getMonth() + 1);
  }

  return months;
}

/**
* 得到一个月的首尾日期
* @param {string} monthString
*/
function getFirstAndLastDayOfMonth(monthString) {
  const [year, month] = monthString.split('/').map(Number);
  const lastDay = new Date(year, month, 0);

  const fullMonth = `${String(month).padStart(2, '0')}`;
  const formattedFirstDay = `${year}/${fullMonth}/01`;
  const formattedLastDay = `${year}/${fullMonth}/${String(lastDay.getDate()).padStart(2, '0')}`;

  return { firstDay: formattedFirstDay, lastDay: formattedLastDay };
}

/**
* 得到最近日期
* @param {number} days
*/
function getRecentDays(days) {
  // 获取当前日期
  let today = new Date();

  let recentDays = [];
  if (days > 0) {
    for (let i = days - 1; i >= 0; i--) {
      let date = new Date(today);
      date.setDate(today.getDate() - i);

      // 格式化日期为 "YYYY/MM/DD" 格式
      let formatDate = function (date) {
        let year = date.getFullYear();
        let month = date.getMonth() + 1;
        let day = date.getDate();
        // 在月份和日期前补0
        const fullMonth = `${String(month).padStart(2, '0')}`;
        const fullDay = `${String(day).padStart(2, '0')}`;

        return year + '/' + fullMonth + '/' + fullDay;
      }(date);

      recentDays.push(formatDate);
    }
  }

  return recentDays;
}

/**
* 深度拷贝
* @param {{ [x: string]: any; mode?: string; criteria?: { field: string; op: string; values: string[]; }[]; hasOwnProperty?: any; }} obj
*/
function deepClone(obj) {
  let objClone = Array.isArray(obj) ? [] : {};
  if (obj && typeof obj === "object") {
    for (let key in obj) {
      if (obj.hasOwnProperty(key)) {
        // 判断obj子元素是否为对象，如果是，递归复制
        if (obj[key] && typeof obj[key] === "object") {
          objClone[key] = deepClone(obj[key]);
        } else {
          // 如果不是，简单复制
          objClone[key] = obj[key];
        }
      }
    }
  }
  return objClone;
}

/**
* 排除相交的记录
* @param {any[]} all
* @param {any[]} excluded
*/
function excludeSomeRecord(all, excluded) {
  return all.filter(obj1 => !excluded.some(obj2 => obj2.id === obj1.id));
}

/**
 * 三天结案率计算
 * 公式=一个月三天(工作日)结案的外部单/一个月登记的外部单(排除最近三天内登记未结案的外部单)
 */
function calcThreeDaysCompleteRate() {
  // 筛选出 三天内登记未结案的外部单(不纳入计算的数据)
  let days = getRecentDays(3);
  let excludedFilter = {
    mode: 'AND',
    criteria: [
      {
        field: '来源',
        op: 'Equals',
        values: ['外部']
      },
      {
        field: '状态',
        op: 'Intersected',
        values: ['', '未开始', '处理中']
      },
      {
        field: '登记日期',
        op: 'Intersected',
        values: [days[0], days[1], days[2]]
      }
    ]
  };
  console.log('excludedFilter: ' + JSON.stringify(excludedFilter));
  console.log('三天内登记未结案的外部单');
  let excluded = fetchAll(origSheet, excludedFilter);
  // excluded.forEach(item => console.log(JSON.stringify(item)));

  // 筛选出 已登记的外部单
  let filter = {
    mode: 'AND',
    criteria: [
      {
        field: '来源',
        op: 'Equals',
        values: ['外部']
      }
    ]
  };
  console.log('所有已登记的外部单');
  let all = fetchAll(origSheet, filter);

  console.log('-------------------------------------------------------------------------------------------');

  if (all && all.length > 0) {

    // 使用sort方法对records数组进行排序
    let sortField = '登记日期';
    sortRecordByField(all, sortField, true);
    // all.forEach(item => console.log(item.fields[sortField]));

    // 登记日期最早的记录
    let startDate = all[0].fields[sortField];
    let endDate = new Date();
    console.log('startDate: ' + startDate + ' endDate: ' + endDate);

    // 获取所有月份
    let months = getMonthsBetweenDates(startDate, endDate);

    months.forEach(function (month) {
      let dateDesc = month.dateDesc;
      console.log('----------------------月份: ' + dateDesc + '----------------------');

      // 一个月的开始和结束
      const { firstDay, lastDay } = getFirstAndLastDayOfMonth(month.dateStr);

      // 筛选出 一个月登记的外部单
      let month_filter = deepClone(filter);
      month_filter.criteria.push({ field: sortField, op: 'GreaterEquAndLessEqu', values: [firstDay, lastDay] });
      console.log('month_filter: ' + JSON.stringify(month_filter));
      console.log(dateDesc + ' 一个月登记的外部单');
      let month_all = fetchAll(origSheet, month_filter);
      month_all = excludeSomeRecord(month_all, excluded);
      console.log('排除不纳入计算的 month_all: ' + month_all.length);

      // 筛选出 一个月三天结案的外部单
      let month_invalidFilter = deepClone(month_filter);
      month_invalidFilter.criteria.push({ field: '状态', op: 'Intersected', values: ['已完成', '已转交'] });
      month_invalidFilter.criteria.push({ field: '结案天数', op: 'GreaterEqu', values: ['0'] });
      month_invalidFilter.criteria.push({ field: '结案天数', op: 'LessEqu', values: ['3'] });
      console.log('month_invalidFilter: ' + JSON.stringify(month_invalidFilter));
      console.log(dateDesc + ' 一个月三天结案的外部单');
      let month_finish = fetchAll(origSheet, month_invalidFilter);

      // 三天结案率
      let rate = month_all.length === 0 ? '' : (month_finish.length / month_all.length).toFixed(4);
      console.log('rate: ' + rate);

      console.log('-------------------------------------------------------------------------------------------');

      // 筛选待更新的指标行
      let rateDesc = '三天结案率';
      let updateFilter = {
        mode: 'AND',
        criteria: [
          {
            field: '月份',
            op: 'Equals',
            values: [dateDesc]
          }
        ]
      };
      console.log('updateFilter: ' + JSON.stringify(updateFilter));
      let updateRows = fetchAll(targetSheet, updateFilter);
      updateRows.forEach(item => console.log(JSON.stringify(item)));

      // 更新指标
      if (updateRows && updateRows.length > 0) {
        updateRows[0].fields[rateDesc] = rate;
        targetSheet.Record.UpdateRecords({
          Records: updateRows
        });
      }
      else {
        targetSheet.Record.CreateRecords({
          Records: [{
            fields: {
              '月份': dateDesc,
              '三天结案率': rate
            }
          }]
        });
      }
    });
  }
}

/**
 * 无效BUG率计算
 * 公式=一个月登记并已分工的无效外部单/一个月登记并已分工的外部单
 */
function calcInvalidBugRate() {
  // 筛选出 已登记的外部单
  let filter = {
    mode: 'AND',
    criteria: [
      {
        field: '来源',
        op: 'Equals',
        values: ['外部']
      }
    ]
  };
  console.log('所有已登记的外部单');
  let all = fetchAll(origSheet, filter);

  console.log('-------------------------------------------------------------------------------------------');

  if (all && all.length > 0) {

    // 使用sort方法对records数组进行排序
    let sortField = '登记日期';
    sortRecordByField(all, sortField, true);
    // all.forEach(item => console.log(item.fields[sortField]));

    // 登记日期最早的记录
    let startDate = all[0].fields[sortField];
    let endDate = new Date();
    console.log('startDate: ' + startDate + ' endDate: ' + endDate);

    // 获取所有月份
    let months = getMonthsBetweenDates(startDate, endDate);

    months.forEach(function (month) {
      let dateDesc = month.dateDesc;
      console.log('----------------------月份: ' + dateDesc + '----------------------');

      // 一个月的开始和结束
      const { firstDay, lastDay } = getFirstAndLastDayOfMonth(month.dateStr);

      // 筛选出 一个月登记并已分工的外部单
      let month_filter = deepClone(filter);
      month_filter.criteria.push({ field: '类型', op: 'NotEmpty', values: [] });
      month_filter.criteria.push({ field: sortField, op: 'GreaterEquAndLessEqu', values: [firstDay, lastDay] });
      console.log('month_filter: ' + JSON.stringify(month_filter));
      console.log(dateDesc + ' 一个月登记并已分工的外部单');
      let month_all = fetchAll(origSheet, month_filter);

      // 筛选出 一个月登记并已分工的无效外部单
      let month_invalidFilter = deepClone(filter);
      month_invalidFilter.criteria.push({ field: '类型', op: 'Intersected', values: ['其他(异常/支援等)', '新版本已修复bug'] });
      month_invalidFilter.criteria.push({ field: sortField, op: 'GreaterEquAndLessEqu', values: [firstDay, lastDay] });
      console.log('month_invalidFilter: ' + JSON.stringify(month_invalidFilter));
      console.log(dateDesc + ' 一个月登记并已分工的无效外部单');
      let month_invalid = fetchAll(origSheet, month_invalidFilter);

      // 无效BUG率
      let rate = month_all.length === 0 ? '' : (month_invalid.length / month_all.length).toFixed(4);
      console.log('rate: ' + rate);

      console.log('-------------------------------------------------------------------------------------------');

      // 筛选待更新的指标行
      let rateDesc = '无效BUG率';
      let updateFilter = {
        mode: 'AND',
        criteria: [
          {
            field: '月份',
            op: 'Equals',
            values: [dateDesc]
          }
        ]
      };
      console.log('updateFilter: ' + JSON.stringify(updateFilter));
      let updateRows = fetchAll(targetSheet, updateFilter);
      updateRows.forEach(item => console.log(JSON.stringify(item)));

      // 更新指标
      if (updateRows && updateRows.length > 0) {
        updateRows[0].fields[rateDesc] = rate;
        targetSheet.Record.UpdateRecords({
          Records: updateRows
        });
      }
      else {
        targetSheet.Record.CreateRecords({
          Records: [{
            fields: {
              '月份': dateDesc,
              '无效BUG率': rate
            }
          }]
        });
      }
    });
  }
}

function main() {
  calcThreeDaysCompleteRate();
  calcInvalidBugRate();
}

main();