let updateRows = [
  {
    "fields": {
      "剩余": 6,
      "总积压": 62,
      "月产出/输入比": 0.942857142857143,
      "月份": "2023年06月",
      "输入": 103,
      "输出": 70
    },
    "id": "Mk"
  }
];

// 删除属性 "剩余" 和 "月产出/输入比"
delete updateRows[0].fields['剩余'];
delete updateRows[0].fields['月产出/输入比'];

console.log(JSON.stringify(updateRows[0]));
console.log(updateRows[0].fields['剩余']); // 会输出 undefined
console.log(updateRows[0].fields['月产出/输入比']); // 会输出 undefined
