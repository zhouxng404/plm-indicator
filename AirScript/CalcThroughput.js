const origSheet = Application.Sheets("问题清单");
// console.log(origSheet.Field.GetFields());
const targetSheet = Application.Sheets("月吞吐量");

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
        } else {
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

    return {firstDay: formattedFirstDay, lastDay: formattedLastDay};
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
 * 吞吐量计算
 * 输入=当月登记的所有单据
 * 输出=当月结案的所有单据
 * 总积压=当月及以前登记的单据，在当月底未结案的单据
 */
function calcThroughput() {
    console.log('所有已登记的单据');
    let all = fetchAll(origSheet);

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
            const {firstDay, lastDay} = getFirstAndLastDayOfMonth(month.dateStr);

            // 注意：结案日期 = 状态最后修改日期，需要加上状态条件筛选
            // 筛选出 一个月登记的单据
            let month_filter = {
                mode: 'AND',
                criteria: [
                    {
                        field: '登记日期',
                        op: 'GreaterEquAndLessEqu',
                        values: [firstDay, lastDay]
                    }
                ]
            };
            console.log(dateDesc + ' 一个月登记的单据');
            let month_all = fetchAll(origSheet, month_filter);
            console.log('month_all: ' + month_all.length);

            // 筛选出 一个月结案的单据
            let month_finishFilter = deepClone(month_filter);
            month_finishFilter.criteria.push({field: '状态', op: 'Intersected', values: ['已完成', '已转交']});
            month_finishFilter.criteria.push({
                field: '结案日期',
                op: 'GreaterEquAndLessEqu',
                values: [firstDay, lastDay]
            });
            console.log('month_finishFilter: ' + JSON.stringify(month_finishFilter));
            console.log(dateDesc + ' 一个月结案的单据');
            let month_finish = fetchAll(origSheet, month_finishFilter);

            // 筛选出 当月及以前登记的数据，现在还未结案的
            let overStockFilter1 = {
                mode: 'AND',
                criteria: [
                    {
                        field: '登记日期',
                        op: 'LessEqu',
                        values: [lastDay]
                    },
                    {
                        field: '状态',
                        op: 'Intersected',
                        values: ['', '未开始', '处理中']
                    }
                ]
            };
            // 筛选出 当月及以前登记的单据，当月以后才结案的
            let overStockFilter2 = {
                mode: 'AND',
                criteria: [
                    {
                        field: '登记日期',
                        op: 'LessEqu',
                        values: [lastDay]
                    },
                    {
                        field: '状态',
                        op: 'Intersected',
                        values: ['已完成', '已转交']
                    },
                    {
                        field: '结案日期',
                        op: 'Greater',
                        values: [lastDay]
                    }
                ]
            };
            console.log(dateDesc + ' 当月及以前登记的单据，在当月底未结案的单据');
            let overStock1 = fetchAll(origSheet, overStockFilter1);
            let overStock2 = fetchAll(origSheet, overStockFilter2);
            let overStock = overStock1.length + overStock2.length;

            console.log('-------------------------------------------------------------------------------------------');

            // 筛选待更新的行
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

            // 更新行
            if (updateRows && updateRows.length > 0) {
                updateRows[0].fields['输入'] = month_all.length;
                updateRows[0].fields['输出'] = month_finish.length;
                updateRows[0].fields['总积压'] = overStock;
                updateRows[0].fields['剩余'] = undefined;
                updateRows[0].fields['月产出/输入比'] = undefined;
                console.log('更新行：' + JSON.stringify(updateRows[0]));
                targetSheet.Record.UpdateRecords({
                    Records: updateRows
                });
            } else {
                targetSheet.Record.CreateRecords({
                    Records: [{
                        fields: {
                            '月份': dateDesc,
                            '输入': month_all.length,
                            '输出': month_all.length,
                            '总积压': overStock
                        }
                    }]
                });
            }
        });
    }
}

function main() {
    calcThroughput();
}

main();