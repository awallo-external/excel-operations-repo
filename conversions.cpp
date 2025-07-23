/* conversions.cpp

  - HOW TO CALL FUNCTIONS

1. include the conversions.h library
2. create an instance of the conversions class
3. call member functions of your instance:


conversions conv; // (step 2)

conv.CSV_2_XLSX(csvPath, xlsxPath); // (step 3)

A note for the not-quite late Mr. Jim:

Using other people's code is ok...
... maybe just pretend that QXlsx is part of Qt,
and also pretend that I didn't change anything in QXlsx to allow for
my fuctions to work? Maybe?

The ever hopeful,


Andrew Wallo V

*/

#include "conversions.h"
#include "xlsxdocument.h"
#include <QFile>
#include <QTextStream>
#include <QDebug>
#include <QFileInfo>

// Many functions call similar code
//
//
// Values for width adjustments, data vs. header toggling,
// and feature activating could be enumerated along with
// function calls for all features that are changed to be
// set/run
// ------------------------------------------------------------

using namespace QXlsx;

conversions::conversions()
{
}

bool conversions::CSV_directTransfer_XLSX_polished(const QStringList &csvPaths, const QString &outputPath)
{
    if (csvPaths.isEmpty()) {
        qWarning() << "No CSV files provided.";
        return false;
    }

    QXlsx::Document finalDoc;

    for (const QString &csvPath : csvPaths)
    {
        QFile file(csvPath);
        if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
            qWarning() << "Failed to open CSV file:" << csvPath;
            return false;
        }

        QTextStream in(&file);
        QString sheetName = QFileInfo(csvPath).baseName();
        finalDoc.addSheet(sheetName);
        Worksheet *sheet = dynamic_cast<Worksheet*>(finalDoc.sheet(sheetName));

        if (!sheet) {
            qWarning() << "Failed to create/access sheet:" << sheetName;
            return false;
        }

        int row = 1;
        while (!in.atEnd()) {
            QString line = in.readLine();
            QStringList values = line.split(',');

            for (int col = 0; col < values.size(); ++col) {
                sheet->write(row, col + 1, values.at(col));
            }

            ++row;
        }

        // Apply formatting
        auto dim = sheet->dimension();
        int firstCol = dim.firstColumn();
        int lastCol  = dim.lastColumn();

        // 1. Freeze top row
        sheet->setFreezeTopRow(true);

        // 2. Add filter
        sheet->setAutoFilter(QXlsx::CellRange(1, 1, 1, lastCol));

        // 3. Header styling
        Format headerFormat;
        headerFormat.setPatternBackgroundColor(QColor("#D3D3D3"));
        headerFormat.setFontBold(true);

        for (int col = firstCol; col <= lastCol; ++col) {
            QVariant val = sheet->read(1, col);
            sheet->write(1, col, val, headerFormat);
        }

        // 4. Adjust width per column
        constexpr double minWidth = 8.43;
        for (int col = firstCol; col <= lastCol; ++col) {
            int maxLen = 0;
            for (int r = 1; r <= dim.lastRow(); ++r) {
                QString str = sheet->read(r, col).toString();
                maxLen = qMax(maxLen, str.length());
            }

            double width = qMax(maxLen * 1.05 + 0.4, minWidth);
            sheet->setColumnWidth(QXlsx::CellRange(1, col, dim.lastRow(), col), width);
        }
    }

    // Save merged workbook to output path
    if (!finalDoc.saveAs(outputPath)) {
        qWarning() << "Failed to save merged XLSX to:" << outputPath;
        return false;
    }

    return true;
}

// Appending tool:
bool conversions::append_CSV_2_XLSX_sheets(const QStringList &csvPaths, const QString &xlsxPath)
{
    if (csvPaths.isEmpty()) {
        qWarning() << "No CSV files provided.";
        return false;
    }

    QXlsx::Document xlsx(xlsxPath);

    if (!xlsx.load()) {
        qWarning() << "Failed to load existing XLSX:" << xlsxPath;
        return false;
    }

    // Helper: track used sheet names
    QSet<QString> existingSheets;
    for (const QString &sheetName : xlsx.sheetNames())
        existingSheets.insert(sheetName);

    // Append each CSV as a new worksheet
    for (const QString &csvPath : csvPaths)
    {
        QFile file(csvPath);
        if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
            qWarning() << "Cannot open CSV:" << csvPath;
            return false;
        }

        QString baseName = QFileInfo(csvPath).baseName();
        QString uniqueName = baseName;
        int counter = 1;
        while (existingSheets.contains(uniqueName)) {
            uniqueName = baseName + QString("(%1)").arg(counter++);
        }

        // Add new sheet with unique name
        xlsx.addSheet(uniqueName);
        Worksheet *sheet = dynamic_cast<Worksheet *>(xlsx.sheet(uniqueName));
        if (!sheet) {
            qWarning() << "Could not create sheet:" << uniqueName;
            return false;
        }

        existingSheets.insert(uniqueName); // Register the new sheet name

        // Write CSV content
        QTextStream in(&file);
        int row = 1;
        while (!in.atEnd()) {
            QString line = in.readLine();
            QStringList values = line.split(',');

            for (int col = 0; col < values.size(); ++col) {
                sheet->write(row, col + 1, values[col]);
            }

            ++row;
        }

        // Format utilities
        auto dim = sheet->dimension();
        int firstCol = dim.firstColumn();
        int lastCol  = dim.lastColumn();
        int lastRow  = dim.lastRow();

        // 1. Freeze top row
        sheet->setFreezeTopRow(true);

        // 2. Add filter
        sheet->setAutoFilter(QXlsx::CellRange(1, 1, 1, lastCol));

        // 3. Style header
        Format headerFormat;
        headerFormat.setPatternBackgroundColor(QColor("#D3D3D3"));
        headerFormat.setFontBold(true);
        for (int col = firstCol; col <= lastCol; ++col) {
            QVariant val = sheet->read(1, col);
            sheet->write(1, col, val, headerFormat);
        }

        // 4. Adjust column widths
        constexpr double minWidth = 8.43;
        for (int col = firstCol; col <= lastCol; ++col) {
            int maxLen = 0;
            for (int r = 1; r <= lastRow; ++r) {
                QString str = sheet->read(r, col).toString();
                maxLen = qMax(maxLen, str.length());
            }
            double width = qMax(maxLen * 1.05 + 0.4, minWidth);
            sheet->setColumnWidth(QXlsx::CellRange(1, col, lastRow, col), width);
        }
    }

    // Save to same file path
    if (!xlsx.saveAs(xlsxPath)) {
        qWarning() << "Failed to save updated XLSX:" << xlsxPath;
        return false;
    }

    return true;
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

// CONTROLS BOLD/BACKGROUND COLOR/FREEZEPANE ACTIVATION/FILTER ACTIVATION
//
bool conversions::XLSX_addUtils(const QString &xlsxPath)
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

    // 5) Save changes back to disk
    if (!xlsx.saveAs(xlsxPath)) {
        qWarning() << "Failed to save XLSX after adding filter.";
        return false;
    }

    return true;
}
bool conversions::XLSX_fixWidth(const QString &xlsxPath, bool isDataWidth)
{
    // 1) Load the workbook
    Document xlsx(xlsxPath);
    if (!xlsx.load()) {
        qWarning() << "Failed to open XLSX for resizing:" << xlsxPath;
        return false;
    }

    // 2) Get worksheet
    Worksheet *sheet = dynamic_cast<Worksheet*>(xlsx.currentWorksheet());
    if (!sheet) {
        qWarning() << "No worksheet available in" << xlsxPath;
        return false;
    }

    // 3) Determine data range
    auto dim = sheet->dimension();
    if (!dim.isValid()) {
        qWarning() << "No valid data range found.";
        return false;
    }

    int firstRow = dim.firstRow();
    int lastRow  = dim.lastRow();
    int firstCol = dim.firstColumn();
    int lastCol  = dim.lastColumn();

    constexpr double minWidth = 8.43;

    // 4) Loop through columns and find max string length
    for (int col = firstCol; col <= lastCol; ++col) {
        int maxLen = 0;

        // Skip header row if isDataWidth is true
        int startRow = isDataWidth ? (firstRow + 1) : firstRow;

        for (int row = startRow; row <= lastRow; ++row) {
            QString str = sheet->read(row, col).toString();
            maxLen = qMax(maxLen, str.length());
        }

        // Width tuning — adjust as needed
        double adjustedWidth = maxLen * 1.05 + 0.4;

        // Enforce minimum width to avoid shrinking default Excel width
        adjustedWidth = qMax(adjustedWidth, minWidth);

        // Apply width to the full range of that column
        sheet->setColumnWidth(CellRange(firstRow, col, lastRow, col), adjustedWidth);
    }

    // 5) Save changes
    if (!xlsx.saveAs(xlsxPath)) {
        qWarning() << "Failed to save adjusted width XLSX:" << xlsxPath;
        return false;
    }

    return true;
}


