

function getEnteredWorkItemCells()
{
  let last_work_item_row = getLastRowOfColumn(input_sheet,INPUT_WORK_ITEM_COLUMN);
  let last_amount_row = getLastRowOfColumn(input_sheet,INPUT_AMOUNT_COLUMN);
  let last_row_index = (last_work_item_row > last_amount_row) ? last_work_item_row : last_amount_row;

  if(last_row_index < INPUT_1ST_ITEM_ROW)
  {
    ui.alert("💡","Please enter one or more work items.\n\nPartial matches work," +
      " using upper or lower case. Looking up \"gasket\" can return \"Head gasket job\" for example.\n\n" +
      CELL_INPUT_WARNING, ui.ButtonSet.OK);
    return {entered_work_items: null, is_blank: true};
  }

  let items = input_sheet.getRange(INPUT_1ST_ITEM_ROW,INPUT_WORK_ITEM_COLUMN,
    last_row_index + 1 - INPUT_1ST_ITEM_ROW);

  return {entered_work_items: items, is_blank: false};
}

function checkPriceList()
{
  let {entered_work_items, is_blank} = getEnteredWorkItemCells();
  
  if(is_blank)
  {
    return false;
  }

  let all_entered_items_have_prices = true;

  // Check all the entered items, skipping blank cells
  for(let row = 1; row <= entered_work_items.getHeight(); row++)
  {
    let entered_desc_cell = entered_work_items.getCell(row,1);
    let neighoring_price_cell = input_sheet.getRange(entered_desc_cell.getRow(),INPUT_AMOUNT_COLUMN);

    if(entered_desc_cell.getValue() === "")
    {
      entered_desc_cell.setFontColor("black");
      neighoring_price_cell.setValue("");
      continue;
    }

    let {price, desc} = lookUpPrice(entered_desc_cell);
    if(desc === null)
    {
      entered_work_items.getCell(row,1).setFontColor(PROBLEM_COLOR);
      neighoring_price_cell.setValue("");
      all_entered_items_have_prices = handleNewOrBadWorkItem(entered_desc_cell);
    }
    else
    {
      entered_desc_cell.setFontColor(SUCCESS_COLOR); // Green
      entered_desc_cell.setValue(desc);
      input_sheet.getRange(entered_desc_cell.getRow(),INPUT_AMOUNT_COLUMN).setValue(price);
    }
  }

  return all_entered_items_have_prices;
}

function handleNewOrBadWorkItem(entered_desc_cell)
{
  let neighoring_price_cell = input_sheet.getRange(entered_desc_cell.getRow(),INPUT_AMOUNT_COLUMN);

  let user_choice = ui.alert("Item not found",
  `\"${entered_desc_cell.getValue()}\" is not in the price table. Would you like to set the price now for this item?`,ui.ButtonSet.YES_NO);

  if(user_choice === ui.Button.YES)
  {
    let response;

    // Prompt until there is a valid number entered
    do
    {
      response = ui.prompt(`Add \"${entered_desc_cell.getValue()}\" to price table`,"$",ui.ButtonSet.OK_CANCEL);

      if(response.getSelectedButton() === ui.Button.CANCEL)
      {
        return false; // Negative result - there's still an invalid entry
      }
    } while(isNaN(Number(response.getResponseText())) || response.getResponseText() === "");
    

    // Add the item and price to the table
    let ps_blank_row = ps_sheet.getLastRow() + 1;
    ps_sheet.getRange(ps_blank_row,PARTS_SERV_DESC_COLUMN).setValue(entered_desc_cell.getValue());
    ps_sheet.getRange(ps_blank_row,PARTS_SERV_PRICE_COLUMN).setValue(Number(response.getResponseText()));

    // Fix the user input and price preview
    entered_desc_cell.setFontColor(SUCCESS_COLOR); // Green
    neighoring_price_cell.setValue(Number(response.getResponseText()));

    return true; // Positive result - added a new entry
  }

  return false; // Negative result - there's still an invalid entry
}

function lookUpPrice(description_cell)
{
  let ps_list = ps_sheet.getRange(PARTS_SERV_1ST_DATA_ROW,PARTS_SERV_DESC_COLUMN,
    ps_sheet.getLastRow() + 1 - PARTS_SERV_1ST_DATA_ROW);
  let result = ps_list.createTextFinder(description_cell.getValue()).findNext();

  if(result === null)
  {
    return {price: null, desc: null};
  }
  else
  {
    // Return the price and the search result
    let price_cell = ps_sheet.getRange(result.getRow(),PARTS_SERV_PRICE_COLUMN);
    return {price: price_cell.getValue(), desc: result.getValue()};
  }
}

function createWorkOrder()
{
  // Store values from input cells into these variables
  let name     = input_sheet.getRange(INPUT_NAME_ROW    ,INPUT_CUSTOMER_COLUMN);
  let phone    = input_sheet.getRange(INPUT_PHONE_ROW   ,INPUT_CUSTOMER_COLUMN);
  let address  = input_sheet.getRange(INPUT_ADDRESS_ROW ,INPUT_CUSTOMER_COLUMN);
  let year     = input_sheet.getRange(INPUT_YEAR_ROW    ,INPUT_CUSTOMER_COLUMN);
  let model    = input_sheet.getRange(INPUT_MODEL_ROW   ,INPUT_CUSTOMER_COLUMN);
  let trim     = input_sheet.getRange(INPUT_TRIM_ROW    ,INPUT_CUSTOMER_COLUMN);
  let plate    = input_sheet.getRange(INPUT_PLATE_ROW   ,INPUT_CUSTOMER_COLUMN);
  let odometer = input_sheet.getRange(INPUT_ODOMETER_ROW,INPUT_CUSTOMER_COLUMN);
  let vin      = input_sheet.getRange(INPUT_VIN_ROW     ,INPUT_CUSTOMER_COLUMN);
  let mechanic = input_sheet.getRange(INPUT_MECHANIC_ROW,INPUT_CUSTOMER_COLUMN);

  if(name.getValue() === "" ||
    model.getValue() === "" ||
    year.getValue() === "" ||
    mechanic.getValue() === "")
  {
    ui.alert("Please search for or enter the customer's " +
      "information and the mechanic doing this job.\n\n" + CELL_INPUT_WARNING);
    return;
  }

  if(!checkPriceList())
  {
    ui.alert("⚠️","Did not create work order. Please correct the entered work items and try again.",ui.ButtonSet.OK);
    return;
  }

  checkForDuplicateWorkOrder(name.getValue(),year.getValue(),model.getValue());

  // Single-row data first
  let blank_wo_row = (wo_sheet.getLastRow() + 1);
  let id_cell = wo_sheet.getRange(blank_wo_row,WO_ID_COLUMN);
  id_cell.setValue(1 + Number(id_cell.getNextDataCell(SpreadsheetApp.Direction.UP).getValue()));

  wo_sheet.getRange(blank_wo_row,WO_DATE_COLUMN).setValue(
    new Date().toLocaleString('en-US', {timeZone: 'America/Chicago'}));
  name    .copyTo(wo_sheet.getRange(blank_wo_row,WO_NAME_COLUMN), {contentsOnly:true});
  year    .copyTo(wo_sheet.getRange(blank_wo_row,WO_YEAR_COLUMN), {contentsOnly:true});
  model   .copyTo(wo_sheet.getRange(blank_wo_row,WO_MODEL_COLUMN), {contentsOnly:true});
  trim    .copyTo(wo_sheet.getRange(blank_wo_row,WO_TRIM_COLUMN), {contentsOnly:true});
  plate   .copyTo(wo_sheet.getRange(blank_wo_row,WO_PLATE_COLUMN), {contentsOnly:true});
  odometer.copyTo(wo_sheet.getRange(blank_wo_row,WO_ODOMETER_COLUMN), {contentsOnly:true});
  mechanic.copyTo(wo_sheet.getRange(blank_wo_row,WO_MECHANIC_COLUMN), {contentsOnly:true});
  phone   .copyTo(wo_sheet.getRange(blank_wo_row,WO_PHONE_COLUMN), {contentsOnly:true});
  address .copyTo(wo_sheet.getRange(blank_wo_row,WO_ADDRESS_COLUMN), {contentsOnly:true});
  vin     .copyTo(wo_sheet.getRange(blank_wo_row,WO_VIN_COLUMN), {contentsOnly:true});

  // Now for multiple row data  
  let {entered_work_items, is_blank} = getEnteredWorkItemCells();
  let item_count = entered_work_items.getHeight();
  let row = blank_wo_row;
  for(; row < blank_wo_row + item_count; row++)
  {
    input_sheet.getRange((row + INPUT_1ST_ITEM_ROW - blank_wo_row),INPUT_WORK_ITEM_COLUMN,1,2).copyTo(
      wo_sheet.getRange(row,WO_WORK_DESC_COLUMN), {contentsOnly:true});
    
    let wo_amount_cell = wo_sheet.getRange(row,WO_AMOUNT_COLUMN);
    let tax_cell = wo_sheet.getRange(row,WO_TAX_COLUMN);
    let subtotal_cell = wo_sheet.getRange(row,WO_SUBTOTAL_COLUMN);

    tax_cell.setValue(`=${wo_amount_cell.getA1Notation()}*${IOWA_TAX}`);
    subtotal_cell.setValue(`=${wo_amount_cell.getA1Notation()}+${tax_cell.getA1Notation()}`);
  }
  row--; // Subtract one to point to the last row with data

  let subtotals = wo_sheet.getRange(blank_wo_row,WO_SUBTOTAL_COLUMN,item_count);
  wo_sheet.getRange(row,WO_TOTAL_COLUMN).setValue(`=sum(${subtotals.getA1Notation()})`);

  // Draw a border line after the work order
  wo_sheet.getRange(`${row}:${row}`).setBorder(null, false, true, false, false, false);

  // Add this info to the list if it's not already there
  updateCustomerList(false,name,phone,address,year,model,trim,plate,vin);

  clearInputs();

  copyDataToInvoice();

  inv_sheet.activate();
}

function checkForDuplicateWorkOrder(name,year,model)
{
  let row_index = getLastRowOfColumn(wo_sheet,WO_ID_COLUMN);

  let last_wo_id = wo_sheet.getRange(row_index,WO_ID_COLUMN).getValue();
  let last_wo_date = wo_sheet.getRange(row_index,WO_DATE_COLUMN).getValue();
  let last_wo_name = wo_sheet.getRange(row_index,WO_NAME_COLUMN).getValue();
  let last_wo_year = wo_sheet.getRange(row_index,WO_YEAR_COLUMN).getValue();
  let last_wo_model = wo_sheet.getRange(row_index,WO_MODEL_COLUMN).getValue();
  
  let day_matches_last_WO = new Date().toDateString() === new Date(last_wo_date).toDateString();
  let name_matches_last_WO = name === last_wo_name;
  let year_matches_last_WO = year === last_wo_year;
  let model_matches_last_WO = model === last_wo_model;

  let user_choice = false;
  let last_wo_work_items = wo_sheet.getRange(row_index,WO_WORK_DESC_COLUMN,
    (1 + wo_sheet.getLastRow() - row_index));

  if(day_matches_last_WO && name_matches_last_WO &&
    year_matches_last_WO && model_matches_last_WO)
  {
    user_choice = promptAboutDuplicateWorkOrder(
      last_wo_id,last_wo_name,last_wo_year,last_wo_model,last_wo_work_items);

    if(user_choice === ui.Button.YES)
    {
      clearLastWorkOrder(true);
    }
  }
}

function clearLastWorkOrder(called_from_function)
{
  if(!called_from_function)
  {
    let row_index = getLastRowOfColumn(wo_sheet,WO_ID_COLUMN);

    let last_wo_id = wo_sheet.getRange(row_index,WO_ID_COLUMN).getValue();
    let last_wo_date = wo_sheet.getRange(row_index,WO_DATE_COLUMN).getValue();
    let last_wo_name = wo_sheet.getRange(row_index,WO_NAME_COLUMN).getValue();
    let last_wo_year = wo_sheet.getRange(row_index,WO_YEAR_COLUMN).getValue();
    let last_wo_model = wo_sheet.getRange(row_index,WO_MODEL_COLUMN).getValue();

    let last_wo_work_items = wo_sheet.getRange(row_index,WO_WORK_DESC_COLUMN,
      (1 + wo_sheet.getLastRow() - row_index));
    let items = last_wo_work_items.getCell(1,1).getValue();

    for(let i = 2; i <= last_wo_work_items.getHeight(); i++)
    {
      items += "\n" + last_wo_work_items.getCell(i,1).getValue();
    }

    let user_choice = ui.alert(`⚠️ Delete work order #${last_wo_id}?`,
      `${last_wo_date}\n\n${last_wo_name}\n${last_wo_year} ${last_wo_model}\n\n${items}`,
      ui.ButtonSet.YES_NO);

    if(user_choice !== ui.Button.YES || last_wo_id === "0")
    {
      return;
    }
  }

  let first_row_of_last_wo = getLastRowOfColumn(wo_sheet,WO_ID_COLUMN);
  let wo_a1_notation = `${first_row_of_last_wo}:${wo_sheet.getLastRow()}`;

  wo_sheet.getRange(wo_a1_notation).clearContent().setBorder(null,false,false,false,false,false);

  let notify_string = called_from_function ? "replaced" : "deleted";
  ui.alert(`Work order ${notify_string}.`);
}

function promptAboutDuplicateWorkOrder(id,name,year,model,work_items)
{
  let line_string = "___________________________________________________________________";
  let items = work_items.getCell(1,1).getValue();

  for(let i = 2; i <= work_items.getHeight(); i++)
  {
    items += "\n" + work_items.getCell(i,1).getValue();
  }

  return ui.alert(`💡 Replace existing work order #${id}?`,
    "A possible duplicate was found in today's records:\n" +
    `${line_string}\n\n${name}\n${year} ${model}\n\n${items}\n${line_string}\n\n` +
    "Would you like to replace this work order with the entered information?", ui.ButtonSet.YES_NO);
}

function copyDataToInvoice()
{
  clearInvoiceFields();

  let row_index = getLastRowOfColumn(wo_sheet,WO_ID_COLUMN);

  let last_wo_id = wo_sheet.getRange(row_index,WO_ID_COLUMN).getValue();
  let last_wo_date     = wo_sheet.getRange(row_index,WO_DATE_COLUMN).getValue();
  let last_wo_name     = wo_sheet.getRange(row_index,WO_NAME_COLUMN).getValue();
  let last_wo_year     = wo_sheet.getRange(row_index,WO_YEAR_COLUMN).getValue();
  let last_wo_model    = wo_sheet.getRange(row_index,WO_MODEL_COLUMN).getValue();
  let last_wo_trim     = wo_sheet.getRange(row_index,WO_TRIM_COLUMN).getValue();
  let last_wo_plate    = wo_sheet.getRange(row_index,WO_PLATE_COLUMN).getValue();
  let last_wo_odometer = wo_sheet.getRange(row_index,WO_ODOMETER_COLUMN).getValue();
  let last_wo_mechanic = wo_sheet.getRange(row_index,WO_MECHANIC_COLUMN).getValue();
  let last_wo_phone    = wo_sheet.getRange(row_index,WO_PHONE_COLUMN).getValue();
  let last_wo_address  = wo_sheet.getRange(row_index,WO_ADDRESS_COLUMN).getValue();
  let last_wo_vin      = wo_sheet.getRange(row_index,WO_VIN_COLUMN).getValue();

  let last_wo_row_count = 1 + wo_sheet.getLastRow() - row_index;
  let last_wo_work_items = wo_sheet.getRange(row_index,WO_WORK_DESC_COLUMN, last_wo_row_count);
  let last_wo_amounts    = wo_sheet.getRange(row_index,WO_AMOUNT_COLUMN, last_wo_row_count);
  let last_wo_taxes      = wo_sheet.getRange(row_index,WO_TAX_COLUMN, last_wo_row_count);

  inv_sheet.getRange(INVOICE_NAME_CELL).setValue(last_wo_name);
  inv_sheet.getRange(INVOICE_ADDRESS_CELL).setValue(last_wo_address);
  inv_sheet.getRange(INVOICE_PHONE_CELL).setValue(last_wo_phone);

  inv_sheet.getRange(INVOICE_VEHICLE_CELL).setValue(`${last_wo_year} ${last_wo_model} ${last_wo_trim}`);
  inv_sheet.getRange(INVOICE_PLATE_CELL).setValue(last_wo_plate);
  if(last_wo_odometer !== "")
  {
    inv_sheet.getRange(INVOICE_MILES_CELL).setValue(`${last_wo_odometer} miles`);
  }
  inv_sheet.getRange(INVOICE_VIN_CELL).setValue(last_wo_vin);

  inv_sheet.getRange(INVOICE_DATE_CELL).setValue(last_wo_date);
  inv_sheet.getRange(INVOICE_ID_CELL).setValue(last_wo_id);
  inv_sheet.getRange(INVOICE_MECHANIC_CELL).setValue(last_wo_mechanic);

  last_wo_work_items.copyTo(inv_sheet.getRange(INVOICE_1ST_ITEM_ROW,INVOICE_WORK_ITEM_COL), {contentsOnly:true});
  last_wo_amounts   .copyTo(inv_sheet.getRange(INVOICE_1ST_ITEM_ROW,INVOICE_PRICE_COL), {contentsOnly:true});
  last_wo_taxes     .copyTo(inv_sheet.getRange(INVOICE_1ST_ITEM_ROW,INVOICE_TAX_COL), {contentsOnly:true});
}

function clearInvoiceFields()
{
  inv_sheet.getRange(INVOICE_NAME_CELL).clearContent();
  inv_sheet.getRange(INVOICE_ADDRESS_CELL).clearContent();
  inv_sheet.getRange(INVOICE_PHONE_CELL).clearContent();

  inv_sheet.getRange(INVOICE_VEHICLE_CELL).clearContent();
  inv_sheet.getRange(INVOICE_PLATE_CELL).clearContent();
  inv_sheet.getRange(INVOICE_MILES_CELL).clearContent();
  inv_sheet.getRange(INVOICE_VIN_CELL).clearContent();

  inv_sheet.getRange(INVOICE_DATE_CELL).clearContent();
  inv_sheet.getRange(INVOICE_ID_CELL).clearContent();
  inv_sheet.getRange(INVOICE_MECHANIC_CELL).clearContent();

  let number_of_columns = 1 + INVOICE_TAX_COL - INVOICE_WORK_ITEM_COL;
  inv_sheet.getRange(INVOICE_1ST_ITEM_ROW,INVOICE_WORK_ITEM_COL,
    INVOICE_LINE_ITEMS,number_of_columns).clearContent();
}
