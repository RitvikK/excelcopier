

public class ExcelCopier {


    /*
     * Get sheet
     * Get list of sheets
     * Get list of rows and columns ( 100 x 50 )
     * Copy from one and set to another
     *
     * */

    XL_ReadWrite xlobj;
    XL_ReadWrite xlCopyobj;

    int iRangeofColsToCopy=10;
    int iRangeofRowsToCopy=10;

    ExcelCopier()
    {
        xlobj=new XL_ReadWrite(FrmConstants.DATAXLSPATH);
        xlCopyobj = new XL_ReadWrite(FrmConstants.COPYFILEPATH);
    }

    private void getSheetDetails()
    {
        printMessages(xlobj.getTotalSheetNum() + " Number of Sheets");
    }

    private void getSheetNamesbyIndex()
    {
        for (int icnt=0;icnt<xlobj.getTotalSheetNum();icnt++)
        {
            printMessages("Sheet " + (icnt+1) + " Name " +  xlobj.getSheetNamebyIndex(icnt));
        }
    }

    private void removeExistingSheets(XL_ReadWrite xlremoveSheets)
    {
        for ( int icnt=0 ; icnt<xlremoveSheets.getTotalSheetNum();icnt++)
        {
            xlremoveSheets.removeSheet(xlremoveSheets.getSheetNamebyIndex(icnt));
        }
        printMessages( "Removed All Sheets Completed.");
    }

    private void addSheets(XL_ReadWrite xlcopySheets,XL_ReadWrite xlbaseSheets)
    {
        for ( int icnt=0 ; icnt<xlbaseSheets.getTotalSheetNum();icnt++)
        {
            xlcopySheets.addSheet(xlbaseSheets.getSheetNamebyIndex(icnt));
        }
        printMessages( "Add Sheets Completed.");
    }

    private void copyDataforSheets(XL_ReadWrite xlcopySheets,XL_ReadWrite xlbaseSheets)
    {
        String copySheetName="";
        printMessages( "Copying Ranges of Rows =" + iRangeofRowsToCopy + "\n Cols =" + iRangeofColsToCopy );
        for (int icnt=0 ; icnt<xlbaseSheets.getTotalSheetNum();icnt++)
        {
            copySheetName=xlbaseSheets.getSheetNamebyIndex(icnt);
            for(int iRow=1; iRow<=iRangeofRowsToCopy;iRow++)
            {
                for (int iCol=0; iCol<iRangeofColsToCopy;iCol++)
                {
                    xlcopySheets.setCellData(copySheetName, iCol, iRow, xlbaseSheets.getCellData(copySheetName, iCol, iRow));
                }
            }

        }
        printMessages( "Copy Data of Sheets Completed.");
    }

    private void printMessages(String stMessage)
    {
        System.out.println(stMessage);
    }

    public static void main(String[] args) {

        ExcelCopier xlfile = new ExcelCopier();
        xlfile.getSheetDetails();
        xlfile.getSheetNamesbyIndex();
        xlfile.removeExistingSheets(xlfile.xlCopyobj);
        xlfile.addSheets(xlfile.xlCopyobj,xlfile.xlobj);
        xlfile.copyDataforSheets(xlfile.xlCopyobj,xlfile.xlobj);


    }

}
