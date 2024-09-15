/**
   * Prints the received response from the socket to MS Excel
   * Texts are printed from the current cursor position
   * Prints only the first result from the response
   * As the first response is the best prediction
   * @param {string} text
   */
export const printInExcel = async (text) => {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const selection = context.workbook.getSelectedRange();

      selection.load("values");
      await context.sync();

      const currentValue = selection.values[0][0];
      const additionalText = text;
      const updatedValue = currentValue + " " + additionalText;

      selection.values = [[updatedValue]];

      await context.sync();
    }).catch(function (error) {
      console.error(error);
    });
  };