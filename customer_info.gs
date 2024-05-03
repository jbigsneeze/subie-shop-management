

function fetchCustomerInfo()
{
  let scriptProperties = PropertiesService.getScriptProperties();
  let customer_index = Number(scriptProperties.getProperty('INDEX')); // Retrieve the stored index
  
  let entered_name = input_sheet.getRange(INPUT_NAME_ROW,INPUT_CUSTOMER_COLUMN).getValue().trim();

  if(entered_name === "")
  {
    ui.alert("💡","Search by first name, last name, or both.\n\nPartial matches work," +
      " using upper or lower case. Looking up \"will\" can return \"William\" for example." +
      "\n\n Repeated button presses will cycle through multiple cars owned by a customer.\n\n" +
      CELL_INPUT_WARNING, ui.ButtonSet.OK);
    return;
  }

  let name_column = cust_sheet.getRange(
    CUSTOMER_1ST_DATA_ROW,CUSTOMER_NAME_COLUMN,cust_sheet.getLastRow() + 1 - CUSTOMER_1ST_DATA_ROW);
  let finder = name_column.createTextFinder(entered_name);

  let found_names = finder.findAll();
  if(found_names.length === 0)
  {
    ui.alert("Customer not found");
    return;
  }

  // If we've gone past the end, cycle over to the start again
  if(customer_index >= found_names.length)
  {
    customer_index = 0;
  }

  let row_index = found_names[customer_index].getRow();
  cust_sheet.getRange(row_index,CUSTOMER_NAME_COLUMN   ).copyTo(input_sheet.getRange(INPUT_NAME_ROW,INPUT_CUSTOMER_COLUMN), {contentsOnly:true});
  cust_sheet.getRange(row_index,CUSTOMER_PHONE_COLUMN  ).copyTo(input_sheet.getRange(INPUT_PHONE_ROW,INPUT_CUSTOMER_COLUMN), {contentsOnly:true});
  cust_sheet.getRange(row_index,CUSTOMER_ADDRESS_COLUMN).copyTo(input_sheet.getRange(INPUT_ADDRESS_ROW,INPUT_CUSTOMER_COLUMN), {contentsOnly:true});
  cust_sheet.getRange(row_index,CUSTOMER_YEAR_COLUMN   ).copyTo(input_sheet.getRange(INPUT_YEAR_ROW,INPUT_CUSTOMER_COLUMN), {contentsOnly:true});
  cust_sheet.getRange(row_index,CUSTOMER_MODEL_COLUMN  ).copyTo(input_sheet.getRange(INPUT_MODEL_ROW,INPUT_CUSTOMER_COLUMN), {contentsOnly:true});
  cust_sheet.getRange(row_index,CUSTOMER_TRIM_COLUMN   ).copyTo(input_sheet.getRange(INPUT_TRIM_ROW,INPUT_CUSTOMER_COLUMN), {contentsOnly:true});
  cust_sheet.getRange(row_index,CUSTOMER_PLATE_COLUMN  ).copyTo(input_sheet.getRange(INPUT_PLATE_ROW,INPUT_CUSTOMER_COLUMN), {contentsOnly:true});
  cust_sheet.getRange(row_index,CUSTOMER_VIN_COLUMN    ).copyTo(input_sheet.getRange(INPUT_VIN_ROW,INPUT_CUSTOMER_COLUMN), {contentsOnly:true});

  scriptProperties.setProperty('INDEX', (customer_index + 1)); // Increment the index and then save it
}

function updateCustomerList(update_existing,name,phone,address,year,model,trim,plate,vin)
{
  let name_column = cust_sheet.getRange(CUSTOMER_1ST_DATA_ROW,CUSTOMER_NAME_COLUMN,
    cust_sheet.getLastRow() + 1 - CUSTOMER_1ST_DATA_ROW);
  let finder = name_column.createTextFinder(name.getValue());

  let found_names = finder.findAll();
  let matching_row = 0;

  // Search for matching entries. Only add an entry if it's unique
  for(let i = 0; i < found_names.length; i++)
  {
    let year_of_found_name = cust_sheet.getRange(found_names[i].getRow(),CUSTOMER_YEAR_COLUMN);
    let model_of_found_name = cust_sheet.getRange(found_names[i].getRow(),CUSTOMER_MODEL_COLUMN);

    let same_year = (year.getValue() === year_of_found_name.getValue());
    let same_model = (model.getValue() === model_of_found_name.getValue());

    if(same_year && same_model)
    {
      // Found a matching record. Don't add a duplicate
      matching_row = model_of_found_name.getRow();
    }
  }

  // Write to list if this is a new record or if we're editing a matching one
  if((matching_row === 0 && !update_existing) ||
    (matching_row !== 0 && update_existing))
  {
    let destination_row = update_existing ? matching_row : cust_sheet.getLastRow() + 1;

    let name_cell = cust_sheet.getRange(destination_row,CUSTOMER_NAME_COLUMN);
    name.copyTo(name_cell, {contentsOnly:true});
    cust_sheet.getRange(destination_row,CUSTOMER_FIRST_NAME_COLUMN).setValue(
      `=split(${name_cell.getA1Notation()}," ")`);

    year   .copyTo(cust_sheet.getRange(destination_row,CUSTOMER_YEAR_COLUMN), {contentsOnly:true});
    model  .copyTo(cust_sheet.getRange(destination_row,CUSTOMER_MODEL_COLUMN), {contentsOnly:true});
    trim   .copyTo(cust_sheet.getRange(destination_row,CUSTOMER_TRIM_COLUMN), {contentsOnly:true});
    plate  .copyTo(cust_sheet.getRange(destination_row,CUSTOMER_PLATE_COLUMN), {contentsOnly:true});
    vin    .copyTo(cust_sheet.getRange(destination_row,CUSTOMER_VIN_COLUMN), {contentsOnly:true});
    phone  .copyTo(cust_sheet.getRange(destination_row,CUSTOMER_PHONE_COLUMN), {contentsOnly:true});
    address.copyTo(cust_sheet.getRange(destination_row,CUSTOMER_ADDRESS_COLUMN), {contentsOnly:true});

    let notify_string = update_existing ? "updated" : "added to customer records";
    ui.alert(`${name.getValue()}'s information has been ${notify_string}.`);
  }
}

function updateCustomerInfo()
{
  // Store values from input cells into these variables
  let name        = input_sheet.getRange(INPUT_NAME_ROW   ,INPUT_CUSTOMER_COLUMN);
  let phone       = input_sheet.getRange(INPUT_PHONE_ROW  ,INPUT_CUSTOMER_COLUMN);
  let address     = input_sheet.getRange(INPUT_ADDRESS_ROW,INPUT_CUSTOMER_COLUMN);
  let year        = input_sheet.getRange(INPUT_YEAR_ROW   ,INPUT_CUSTOMER_COLUMN);
  let model       = input_sheet.getRange(INPUT_MODEL_ROW  ,INPUT_CUSTOMER_COLUMN);
  let trim        = input_sheet.getRange(INPUT_TRIM_ROW   ,INPUT_CUSTOMER_COLUMN);
  let plate       = input_sheet.getRange(INPUT_PLATE_ROW  ,INPUT_CUSTOMER_COLUMN);
  let vin         = input_sheet.getRange(INPUT_VIN_ROW    ,INPUT_CUSTOMER_COLUMN);

  if(name.getValue() === "" ||
    model.getValue() === "" ||
    year.getValue() === "" )
  {
    SpreadsheetApp.getUi().alert("Please enter the customer " +
      "name and the vehicle year and model, along with updated customer information.\n\n" +
      CELL_INPUT_WARNING);
  }
  else
  {
    updateCustomerList(true,name,phone,address,year,model,trim,plate,vin);
  }
}

function clearInputs()
{
  // This clears a few extra rows but whatever
  let num_rows = input_sheet.getLastRow();

  input_sheet.getRange(1,INPUT_CUSTOMER_COLUMN,num_rows).clearContent().setFontColor("black");
  input_sheet.getRange(INPUT_1ST_ITEM_ROW,INPUT_WORK_ITEM_COLUMN,num_rows,2).clearContent().setFontColor("black");
}
