#include "Convertor.h"

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);
    QTextStream qtin(stdin);
    QTextStream qtout(stdout);
     QString f_path;
       string f_std_str_path;
    QString sh_name;
    bool PrintDataFlag=true,PrintNameFlag=true;
    int ch_1,ch_2;
    qtout<<"welcome to xlsx_to_json converter"<<endl;
    qtout << "enter  path to file:" << endl;
    qtout << "example: 'D:\\folder1\\folder2\\file.xlsx'" << endl;
    getline(cin,f_std_str_path) ;
    f_path=Convertor::ParsePath( QString::fromStdString( f_std_str_path));
    qtout<<" want you print names?"<<endl;
    qtout<<"press 0 if not? else press any key"<<endl;
    qtin>>ch_1;
    PrintNameFlag=ch_1;
    qtout<<" want you  also print data to console?"<<endl;
    qtout<<"press 0 if not? else press any key"<<endl;
    qtin>>ch_2;
    PrintDataFlag=ch_2;
    Convertor cnv(f_path,PrintDataFlag,PrintNameFlag);
        cnv.OpenBook();
    QStringList sh_list= cnv.OpenWorkSheet();
    int ch_number=0;
    cout<<"choose current worksheet. enter the number"<<endl;
    cin>>ch_number;
    if (ch_number>=0 && ch_number<= cnv.GetSheetCount() ) {
        cnv.SetActivetWorkSheet(sh_list[ch_number]);
    }
    cnv.GetActivetWorkSheet();
    cnv.SetMaxRows();
    cnv.SetMaxColumns();
    cnv.Generate();


    cout << endl;



    return a.exec();
}

