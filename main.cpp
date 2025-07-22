// main.cpp

#include <QCoreApplication>
#include <QDebug>
#include "conversions.h"
#include "xlsxmerger.h"

int main()
{
    QStringList inputs = { "C:/Users/Andrew.Wallo1/Documents/misc/testing_merger/input1.xlsx",
                            "C:/Users/Andrew.Wallo1/Documents/misc/testing_merger/input2.xlsx",
                             "C:/Users/Andrew.Wallo1/Documents/misc/testing_merger/input3.xlsx"};

    QString output =  "C:/Users/Andrew.Wallo1/Documents/misc/testing_merger/merged_output.xlsx";

    XlsxMerger merger;
    if (!merger.mergeFiles(inputs, output)) {
        qWarning() << "Failed to merge XLSX files.";
    }

    return 0;
}



/* TEST MAIN 1 */

/*
int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    QString csvPath  = QStringLiteral("C:/Users/Andrew.Wallo1/Documents/misc/testing_site/input_1.csv");
    QString xlsxPath = QStringLiteral("C:/Users/Andrew.Wallo1/Documents/misc/testing_site/output2.xlsx");

    conversions conv;

    if (!conv.CSV_2_XLSX(csvPath, xlsxPath)) {
        qCritical() << "CSV → XLSX conversion failed";
        return 1;
    }
    if (!conv.XLSX_addUtils(xlsxPath)) {
        qCritical() << "Adding AutoFilter failed";
        return 1;
    }
    if (!conv.XLSX_fixWidth(xlsxPath, true)) { // TRUE => Only adjusts On data // Switch to FALSE for header adjustments
        qCritical() << "Adding width adjustments failed";
        return 1;
    }
    qDebug() << "Done!";
    return 0;
}

*/

// NEXT GEN development: Individual Column Adjustment selection based on an array/vector of boolean inputs

// Consider selection options for user on dextor side for each function -> rewrite functions for more toggle abilities



