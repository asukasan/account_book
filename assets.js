function load_sheet(name, range=null){;
    let spread_sheet = SpreadsheetApp.getActive()
    let sheet = spread_sheet.getSheetByName(name);
    if(range == null){
      return sheet
    }else{
      let sheet_range = sheet.getRange(range);
      return sheet_range;
    }
  }
  
  function add_transaction(){
    const TRANSACTION_POSITION = "C7:G7";
    let sheet = load_sheet("transaction");
    transaction = sheet.getRange(TRANSACTION_POSITION);
    let id_range = sheet.getRange("B13:B");
    let latest_id_count = id_range.getValues().filter(String).length + 1;
    let new_transaction_position = latest_id_count + 12;
    sheet.getRange("B" + new_transaction_position.toString()).setValue(latest_id_count);
    sheet.getRange("C" + new_transaction_position.toString() + ":" + "G" + new_transaction_position.toString()).setValues(transaction.getValues());
    let account_name = transaction.getValues()[0][4];
    let money = transaction.getValues()[0][2];
    calc_account(account_name, money)
  
    extra_transaction("week_account_sum", money, account_name);
    extra_transaction("month_account_sum", money, account_name);
    if(money < 0){
      // category_transactionを起動する。　
      let category_name = transaction.getValues()[0][3];
      extra_transaction("week_category_sum", money, category_name);
      extra_transaction("month_category_sum", money, category_name);
    }
  
    default_transaction();
  }
  
  function default_transaction(){
    const TRANSACTION_POSITION = "D7:G7"
    let range = load_sheet("transaction", TRANSACTION_POSITION);
    range.clear();
  }
  
  function calc_account(name="現金", transaction_money=300){
    const ASSETS_POSITION = "C3:R3"
    let sheet = load_sheet("category")
    let assets = sheet.getRange(ASSETS_POSITION); 
    var place = assets.createTextFinder(name).findAll();
    let money_position = sheet.getRange(4, place[0].getColumn());
    let account_money = money_position.getValue();
    let sum = account_money + transaction_money;
    money_position.setValue(sum);
  
  }
  
  function week_process(){
    var date = new Date();
    Logger.log(date);
    Logger.log(date.getMonth() + 1);
    Logger.log(date.getDate());
    // 一週間を取得
  }
  
  function feature_process(){
    // 毎月27日に実行
    let today = new Date();
    let max_num;
    let month = today.getMonth() + 1;
    if(month == 1 || month == 7){
      max_num = 4;
    }else{
      max_num = 2;
    }
    for(let i = 0; max_num > i; i++){
      let sheet = load_sheet("feature_pay",);
      let range = sheet.getRange(6 + i, 3, 1, 11);
      let name = range.getValues()[0][2]
      let category = range.getValues()[0][3]
      let money = range.getValues()[0][7]
      let account = range.getValues()[0][4]
      // countを-1する
      let count_position = sheet.getRange(6+i, 9);
      let count = count_position.getValue();
      count_position.setValue(count - 1);
      set_transaction(name, -money, category, account);
      add_transaction();
    }
  }
  
  function sample(){
    set_transaction("hoge", -300, "食費", "sample");
  }
  
  function set_transaction(name, money, category, account){
    // forで一つずつ入れていく
    let sheet = load_sheet("transaction");
    sheet.getRange(7, 4).setValue(name);
    sheet.getRange(7, 5).setValue(money);
    sheet.getRange(7, 6).setValue(category);
    sheet.getRange(7, 7).setValue(account);
  
  }
  
  
  function extra_transaction(sheet_name, money, kind){
    // categoryの場合はmoneyの符号を変える
    if(sheet_name.match(/category/)){
      money = - money;
    }
    let sheet = load_sheet(sheet_name);
    let date_range = load_sheet(sheet_name, "B4:B");
    let kind_range = load_sheet(sheet_name, "C3:Z3");
    let current_date_position = date_range.getValues().filter(String).length + 3;
    var kind_position = kind_range.createTextFinder(kind).findAll();
    // kindの行番号はplace[0].getColumn()で取得できる。
    let set_position = sheet.getRange(current_date_position, kind_position[0].getColumn())
    kind_sum = set_position.getValue();
    let sum = kind_sum + money;
    set_position.setValue(sum);
  
  }
  
  function add_week(){
    // categoryとacctountのページの最終行の位置を取得して、その下の行に今日の日付をsetする。
    let account_sheet = load_sheet("week_account_sum");
    let category_sheet = load_sheet("week_category_sum");
    let account_range = account_sheet.getRange("B4:B");
    let category_range = category_sheet.getRange("B4:B");
    let today = new Date();
    let latest_account_position = account_range.getValues().filter(String).length + 1;
    let latest_category_position = category_range.getValues().filter(String).length + 1;
    let new_latest_account_position = latest_account_position + 3;
    let new_latest_category_position = latest_account_position + 3;
    account_sheet.getRange("B" + new_latest_account_position.toString()).setValue(today);
    category_sheet.getRange("B" + new_latest_category_position.toString()).setValue(today);
  }
  
  function add_month(){
    let account_sheet = load_sheet("month_account_sum");
    let category_sheet = load_sheet("month_category_sum");
    let account_range = account_sheet.getRange("B4:B");
    let category_range = category_sheet.getRange("B4:B");
    let today = new Date();
    let latest_account_position = account_range.getValues().filter(String).length + 1;
    let latest_category_position = category_range.getValues().filter(String).length + 1;
    let new_latest_account_position = latest_account_position + 3;
    let new_latest_category_position = latest_account_position + 3;
    account_sheet.getRange("B" + new_latest_account_position.toString()).setValue(today);
    category_sheet.getRange("B" + new_latest_category_position.toString()).setValue(today);
  }
  