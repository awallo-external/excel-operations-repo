// main.cpp

#include <QCoreApplication>
#include <QDebug>
#include <QDir>
#include <QFileInfo>
#include <QStringList>
#include "conversions.h"
#include "xlsxmerger.h"


// MAIN_CPP is primarily used as a test_bench for calling the functions.



// FUNCTION DECLARATIONS //
QStringList getFilesOfType(const QString &folderPath, const QString &extension);
QStringList getFullFilePathsOfType(const QString &folderPath, const QString &extension);
void complete_actions(); // build-out for TEST 3
// ------------------------------------------------------------------------------



/* TEST MAIN 3: CSV files --> XLSX files */
int main()
{
    complete_actions();
    return 0;
}



// ------------------------------------------------------------------------------
// SOME FUNCITONS (Alternative tests are stored below functions)



// function grabs all of a certain file type from a folder
QStringList getFullFilePathsOfType(const QString &folderPath, const QString &extension)
{
    QDir dir(folderPath);
    if (!dir.exists()) {
        qWarning() << "Directory does not exist:" << folderPath;
        return {};
    }

    QStringList filters;
    filters << "*." + extension;

    QFileInfoList fileInfoList = dir.entryInfoList(filters, QDir::Files);
    QStringList fullPaths;
    for (const QFileInfo &info : fileInfoList) {
        fullPaths << info.absoluteFilePath();
    }
    return fullPaths;
}

// Innards for Test # // complete test for append functions and direct transfer as needed...
void complete_actions(){

    conversions conv;

    // BUILD TEST QStringList for inputting .csv files:
    QStringList inputFiles ={ "C:/Users/Andrew.Wallo1/Documents/misc/test_folders_csv2xlsx/myCSVstuff/input1.csv",
                              "C:/Users/Andrew.Wallo1/Documents/misc/test_folders_csv2xlsx/myCSVstuff/input2.csv",
                              "C:/Users/Andrew.Wallo1/Documents/misc/test_folders_csv2xlsx/myCSVstuff/input3.csv",
                            };


    QString outputFile = "C:/Users/Andrew.Wallo1/Documents/misc/test_folders_csv2xlsx/final_output.xlsx";

   //  conv.CSV_directTransfer_XLSX_polished(inputFiles, outputFile);
    conv.append_CSV_2_XLSX_sheets(inputFiles, outputFile);

}

/* TEST MAIN 2: XLSX COMBINER */

/*

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

*/

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



