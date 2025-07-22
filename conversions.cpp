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


bool conversions::XLSX_addFilter(const QString &xlsxPath)
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
    //    Here we take the sheet's used dimension
    auto dim = sheet->dimension();
    int lastCol = dim.lastColumn();
    if (lastCol < 1) {
        qWarning() << "Worksheet appears empty, nothing to filter.";
        return false;
    }

    // 4) Set the AutoFilter on row 1, from col 1 ("A1") to col lastCol
    sheet->setAutoFilter(CellRange(1, 1, 1, lastCol));

    // 5) Save changes back to disk
    if (!xlsx.saveAs(xlsxPath)) {
        qWarning() << "Failed to save XLSX after adding filter.";
        return false;
    }

    return true;
}

/*
bool conversions::XLSX_addFilter(const QString &xlsxPath)
{
    Document xlsx(xlsxPath);
    Worksheet *sheet = dynamic_cast<Worksheet*>(xlsx.currentWorksheet());
    if(!sheet)
    {
        qWarning() << "No valid worksheet found in file: " << xlsxPath;
    }

    // Find the last column in the first row:
    int colCount = 0;
    for (int col = 1; col < 1000; ++col)
    {
        if (sheet->read(1,col).isValid())
            colCount = col;
        else
            break;
    }

    if (colCount == 0)
    {
        qWarning() << "No header row found to apply filter.";
        return false;
    }

    // Apply filter to first row
    CellRange filterRange(1,1,1, colCount);
    sheet->setAutoFilter(filterRange);

    if (!xlsx.save())
    {
        qWarning() << "Failed to save XLSX file after adding filter.";
        return false;
    }

    return true;
}

*/
