#include <QCoreApplication>
#include <QStringList>
#include <QDebug>

#include "atk.h"  // Your toolkit header

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);

    QStringList args = app.arguments();

    if (args.size() != 3) {
        qCritical() << "Usage: " << args[0] << " <input1.xlsx> <input2.xlsx>";
        return 1;
    }

    // obtain paths from command line
    QString path1 = QDir::fromNativeSeparators(argv[1]);
    QString path2 = QDir::fromNativeSeparators(argv[2]);

    qDebug() << "Input file 1:" << path1;
    qDebug() << "Input file 2:" << path2;

    return 0;
}
