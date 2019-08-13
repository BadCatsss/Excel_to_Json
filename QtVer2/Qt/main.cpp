#include <QCoreApplication>
#include "Convertor.h"

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);
    //Convertor cnv( QString::fromStdString( "D:\\old.xlsx"));//test open
    if (argc>1) {


        Convertor cnv(QCoreApplication::arguments().at(1));
        if (cnv.openBook()) {
            cout << "file was open" << endl;
            auto shList = cnv.getSheetsList();
            for (auto i : shList)
                cout << i.toStdString() << endl;
            cout << "choose sheet" << endl;
            int chooseNumber;
            cin >> chooseNumber;
            if (chooseNumber >= 0 && chooseNumber <= shList.size()) {
                cnv.setActivetWorkSheet(shList[chooseNumber]);
                cout << shList[chooseNumber].toStdString() << endl;
            }
            else {
                cout << "incorrect list number";
            }
            cnv.convert();
            return a.exec();
        }
        else {
            cout << "cant open file" << endl;
            exit(-1);
        }
    }
    else {
        cout << "Program was run whitout arguments" <<endl;
    }
}
