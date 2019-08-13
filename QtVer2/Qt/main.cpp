#include <QCoreApplication>
#include "Convertor.h"

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);
    //   Convertor cnv("D:\\old.xlsx");//test open
    Convertor cnv(argv[1]);
    cnv.convert();
    return a.exec();
}
