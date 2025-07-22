#include "conversions.h"
#include "xlsxdocument.h"

#include <QFile>
#include <QTextStream>
#include <QDebug>

using namespace QXlsx;

conversions::conversions()
{
}

// Converts a Comma Separated File (CSV) into a zipped folder set of XML files (Excel file format)
//
// Function returns true upon successful conversion and output of CSV file in XLSX format
//
bool conversions::CSV_2_XLSX(const QString &csvPath, const QString &xlsxPath)
{
    QFile file(csvPath);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        qWarning() << "Failed to open CSV file:" << csvPath;
        return false;
    }

    QTextStream in(&file);
    Document xlsx;

    int row = 1;
    while (!in.atEnd()) {
        QString line = in.readLine();
        QStringList values = line.split(',');

        for (int col = 0; col < values.size(); ++col) {
            xlsx.write(row, col + 1, values.at(col));
        }

        ++row;
    }

    if (!xlsx.saveAs(xlsxPath)) {
        qWarning() << "Failed to save XLSX file:" << xlsxPath;
        return false;
    }

    return true;
}


bool conversions::XLSX_addUtils(const QString &xlsxPath) // ALSO HANDLES FREEZE PANE
{
    // 1) Load the workbook from disk
    Document xlsx(xlsxPath);
    if (!xlsx.load()) {
        qWarning() << "Failed to open XLSX for editing:" << xlsxPath;
        return false;
    }

    // 2) Grab the first worksheet
    Worksheet *sheet = dynamic_cast<Worksheet*>(xlsx.currentWorksheet());
    if (!sheet) {
        qWarning() << "No worksheet available in" << xlsxPath;
        return false;
    }

    // 3) Determine how many columns your header row spans
    //   take the sheet's used dimension
    auto dim = sheet->dimension();
    int lastCol = dim.lastColumn();
    if (lastCol < 1) {
        qWarning() << "Worksheet appears empty, nothing to filter.";
        return false;
    }


    // 4) Set the AutoFilter on row 1, from col 1 ("A1") to col lastCol
    CellRange filterRange(1, 1, 1, lastCol);
    sheet->setAutoFilter(filterRange);

    // Set the freeze pane instruction as well ...
    sheet->setFreezeTopRow(true);

    // Apply grey background to header row
    Format headerFormat;
    headerFormat.setPatternBackgroundColor(QColor("#D3D3D3")); // Light grey
    headerFormat.setFontBold(true);

    int firstRow = filterRange.firstRow();
    int firstCol = filterRange.firstColumn();
    int lastColumn = filterRange.lastColumn();

    for (int col = firstCol; col <= lastColumn; ++col) {
        QVariant val = sheet->read(firstRow, col);
        sheet->write(firstRow, col, val, headerFormat);
    }

    for (int col = dim.firstColumn(); col <= dim.lastColumn(); ++col) {
        int maxLen = 0;
        for (int row = dim.firstRow(); row <= dim.lastRow(); ++row) {
            QString str = sheet->read(row, col).toString();
            maxLen = qMax(maxLen, str.length());
        }

        // Tighter approximation to avoid oversizing
        double adjustedWidth = maxLen * 1.05 + 0.4; // SET PADDING DOUBLES to 1.2 and 1.0 for wider approximation

        // Corrected: apply column width for just this column
        sheet->setColumnWidth(CellRange(1, col, dim.lastRow(), col), adjustedWidth);
    }

    // 5) Save changes back to disk
    if (!xlsx.saveAs(xlsxPath)) {
        qWarning() << "Failed to save XLSX after adding filter.";
        return false;
    }

    return true;
}

