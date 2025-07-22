#include "xlsxmerger.h"
#include "xlsxdocument.h"
#include "xlsxworksheet.h"

#include <QFileInfo>
#include <QDebug>

using namespace QXlsx;

XlsxMerger::XlsxMerger() {}

bool XlsxMerger::mergeFiles(const QStringList &inputPaths, const QString &outputPath)
{
    Document outputDoc;

    for (const QString &path : inputPaths)
    {
        Document inputDoc(path);
        if (!inputDoc.load()) {
            qWarning() << "Failed to load:" << path;
            return false;
        }

        Worksheet *inputSheet = dynamic_cast<Worksheet*>(inputDoc.currentWorksheet());
        if (!inputSheet) {
            qWarning() << "Invalid worksheet in file:" << path;
            return false;
        }

        QString sheetName = QFileInfo(path).baseName();
        outputDoc.addSheet(sheetName);
        Worksheet *outSheet = dynamic_cast<Worksheet*>(outputDoc.sheet(sheetName));
        if (!outSheet) {
            qWarning() << "Failed to create sheet:" << sheetName;
            return false;
        }

        auto dim = inputSheet->dimension();
        for (int row = dim.firstRow(); row <= dim.lastRow(); ++row) {
            for (int col = dim.firstColumn(); col <= dim.lastColumn(); ++col) {
                QVariant value = inputSheet->read(row, col);
                outSheet->write(row, col, value);
            }
        }
    }

    return outputDoc.saveAs(outputPath);
}
