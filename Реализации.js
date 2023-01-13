const request = require('request');
const { google } = require('googleapis');
const keys = require('./.json')
const token = ''


const clientCup = new google.auth.JWT(
  keys.client_email,
  null,
  keys.private_key,
  ['https://www.googleapis.com/auth/spreadsheets']
);


clientCup.authorize(function (err, tokens) {
  if (err) {
    console.log(err);
    return;
  }
  else {
    console.log('Connected!');
    gsrunCup(clientArrdid);
  }

});

async function gsrunCup(cl) { // Шабка //
  const gsapiCup = google.sheets({ version: 'v4', auth: cl });
  let Cap = [["rd_id", "gi_id", "subject_name", "nm_id", "brand_name", "sa_name", "ts_name", "barcode", "doc_type_name", "quantity", "retail_amount", "office_name", "supplier_oper_name", "order_dt", "sale_dt", "shk_id", "delivery_amount", "return_amount", "delivery_rub", "gi_box_type_name", "rid", "ppvz_sales_commission", "ppvz_for_pay", "ppvz_vw_nds", "sticker_id", "site_country", "penalty", "additional_payment",
    "srid", "bonus_type_name"]] // Список шабки

  const updateOptionsCup = {
    spreadsheetId: '', // Ссылка на таблицу
    range: 'Sheet1!A1', // Ячейка и лист
    valueInputOption: 'USER_ENTERED',
    resource: { values: Cap } // Данные
  };
  let resCup = await gsapiCup.spreadsheets.values.update(updateOptionsCup); // Запись данных в таблицу
} //



const clientArrdid = new google.auth.JWT(
  keys.client_email,
  null,
  keys.private_key,
  ['https://www.googleapis.com/auth/spreadsheets']
);


clientArrdid.authorize(function (err, tokens) {
  if (err) {
    console.log(err);
    return;
  }
  else {
    console.log('Connected!');
    gsrunArrdid(clientArrdid);
  }

});

async function gsrunArrdid(cl) {
  const gsapiArrdid = google.sheets({ version: 'v4', auth: cl });

  const optArrdid = {
    spreadsheetId: '', //Ссылка на таблицу
    range: 'Sheet1!A2:A10000000' //С каких ячеек брать данные
  };
  let dataArrdid = await gsapiArrdid.spreadsheets.values.get(optArrdid);
  let if_else_Arrdid = dataArrdid.data.values // Данные из таблицы

  if(if_else_Arrdid === undefined) { // Если нету данных в таблице
    var lastRow = '2'
  } else { //Если есть данные в таблице
    var lastRow = dataArrdid.data.values.length - 1 //Последния строка

  }

  if (if_else_Arrdid === undefined) { // Если в таблице нет данных
    var arrdid = `&rrdid=0`
    console.log('В таблице нет данных')
  }
  else { // Если в таблице есть данные
    let lastRowArrdid = dataArrdid.data.values.length - 1 //Последния строка
    let data = dataArrdid.data.values[[lastRowArrdid]] // Элимент в страке

    // console.log(data2.data.values[[lastRow]])

    var arrdid = `&rrdid=${data}`
    console.log('Данные в таблице есть')
  }



  // Запрос на сервер
  const headers = { 'Authorization': token, }
  let aDateFrom = "dateFrom=2022-09-01";
  let aDateTo = "&dateTo=2050-03-25";
  let alimit = "&limit=2000"
  request.get({
    url: `https://statistics-api.wildberries.ru/api/v1/supplier/reportDetailByPeriod?${aDateFrom}&key=${alimit}${arrdid}${aDateTo}`,
    headers: headers,
  }, //
    function (error, response, body) {
      let rows = []
      let dataSet = JSON.parse(body) // Делаем body в json



      const client = new google.auth.JWT(
        keys.client_email,
        null,
        keys.private_key,
        ['https://www.googleapis.com/auth/spreadsheets']
      );


      client.authorize(function (err, tokens) {
        if (err) {
          console.log(err);
          return;
        }
        else {
          console.log('Connected!');
          gsrun(client);
        }

      });

      async function gsrun(cl) {
        const gsapi = google.sheets({ version: 'v4', auth: cl });

        // const opt = {
        //   spreadsheetId: '',
        //   range: 'Sheet1!A2:A100000000'
        // };
        // let data = await gsapi.spreadsheets.values.get(opt);

        dataSet.forEach(function (el) { // Проходим по всем элиментам
          let { rrd_id, gi_id, subject_name, nm_id, brand_name, sa_name, ts_name, barcode, doc_type_name, quantity, retail_amount, office_name, supplier_oper_name, order_dt, sale_dt, shk_id, delivery_amount, return_amount, delivery_rub, gi_box_type_name, rid, ppvz_sales_commission, ppvz_for_pay, ppvz_vw_nds, sticker_id, site_country, penalty, additional_payment,
            srid, } = el

          let bonus_type_name = el.bonus_type_name
          rows.push([rrd_id, gi_id, subject_name, nm_id, brand_name, sa_name, ts_name, barcode, doc_type_name, quantity, retail_amount, office_name, supplier_oper_name, order_dt, sale_dt, shk_id, delivery_amount, return_amount, delivery_rub, gi_box_type_name, rid, ppvz_sales_commission, ppvz_for_pay, ppvz_vw_nds, sticker_id, site_country, penalty, additional_payment,
            srid, bonus_type_name])
        })

        const updateOptions = {
          spreadsheetId: '', // Ссылка на таблицу
          range: `Sheet1!A${lastRow}`, // Ячейка и лист
          valueInputOption: 'USER_ENTERED',
          resource: { values: rows } // Данные
        };
        let res = await gsapi.spreadsheets.values.update(updateOptions); // Запись данных в таблицу

      }
    })

}

