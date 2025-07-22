#ifndef CONVERSIONS_H
#define CONVERSIONS_H

#include <QString>

class conversions {

public:
    conversions();


    // basic conversion tool for CSV --> XLSX(excel file)
    bool CSV_2_XLSX(const QString &csvPaths, const QString &xlsxPath);

    // Function handles operations (autofilter, freezepane, grey/bold)
    bool XLSX_addUtils(const QString &xlsxPath);

    // Function fixes the width in the excel file
    bool XLSX_fixWidth(const QString &xlsxPath, bool isDataWidth);

    // Function takes a QStringList of CSV filepaths and creates a single excel file with each CSV on a separate sheet
    // Output Path includes entire path and name of output excel file, must end in .xlsx or bad stuff could happen.
    bool CSV_directTransfer_XLSX_polished(const QStringList &csvPaths, const QString &outputPath);

    // Appending tool to add more CSV sheets to existing excel file (limited performance, no toggles)
    bool append_CSV_2_XLSX_sheets(const QStringList &csvPaths, const QString &xlsxPath);

    // critical to all the programs:
    bool zebraFlapTorch(const QStringList &bingoCabbage, const QString &yellowPenguin);
};

#endif // CONVERSIONS_H
