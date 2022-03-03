#include "widget.h"
#include "ui_widget.h"
#include <QSqlRecord>
#include <QDate>
#include <iostream>
#include <QtPrintSupport/QPrinter>
#include <QtPrintSupport/QPrinterInfo>
#include <QTableWidget>
#include <QDir>
//test
//тест гита 1222
// test git
Widget::Widget(QWidget *parent) : QWidget(parent) , ui(new Ui::Widget)
{
    proc = new QProcess(this);
    ui->setupUi(this);
    db = QSqlDatabase::addDatabase("QSQLITE");
    db.setDatabaseName("./DB_for_amb.db");
    if(db.open()){
        qDebug()<<"Open";
    }
    else {qDebug()<<"Not open";}
    query = new QSqlQuery(db);
    query_lec=NULL;

    connect(ui->fio_button,SIGNAL(clicked()),this,SLOT(FIO_slot()));//коннект кнопки входа врача в систему и рецепта
    connect(ui->searchbutton,SIGNAL(clicked()),this,SLOT(search_slot()));//коннект для поиска в базе данных пациентов
    connect(ui->pacient_button,SIGNAL(clicked()),this,SLOT(choose_slot()));//коннект для выбора пациента, когда все ок он выводит на последний слайд пациента
    connect(ui->addNew_button,SIGNAL(clicked()),this,SLOT(addNewUser_slot()));//коннект для добавления нового пациента
    connect(ui->fill_button,SIGNAL(clicked()),this,SLOT(filling_slot()));//коннект для заполнения документа  exel под конкретного пациента и врача
    connect(ui->deleteUser_button,SIGNAL(clicked()),this,SLOT(delete_user()));//коннект для удаления пользователя
    connect(ui->findAndFillLec_button,SIGNAL(clicked()),this,SLOT(set_lec()));//нахождение лекарства  в exel и выведение его в окно приложения
    connect(ui->findAndFillLec_button,SIGNAL(clicked()),this,SLOT(set_db_lec()));//заполнение базы данных врачей выписанными рецептами
    connect(ui->clear_button,SIGNAL(clicked()),this,SLOT(clear_slot()));//слот очищения всех полей в программе
    connect(ui->fill_button,SIGNAL(clicked()),this,SLOT(print()));//печать документа на принтер, скорее всего будет испольховаться другая кнопка
}

void Widget::FIO_slot(){//слот который устанавливает фио врача на последнюю страницу
    ui->doclabel->setText(ui->lineEdit->text());
}

void Widget::search_slot(){//слот поиска пациента по базе данных пациентов
    open_if_open();
    query = new QSqlQuery(db);
    query->exec("SELECT * FROM Pacient WHERE Firstname = '"+ui->firstnameline->text().split(" ")[1]+"' AND Lastname = '"+ui->firstnameline->text().split(" ")[0]+"' ");
        while(query->next()){
            ui->podrobneelabel->setText(query->value(0).toString()+" "+query->value(1).toString()+" "+query->value(2).toString()+" "+query->value(3).toString());
        }
    if(ui->podrobneelabel->text()!="/*Полная Информация*/"){
        ui->resultlabel->setText("Success");
    }
    else{
        ui->resultlabel->setText("False");
    }
}

void Widget::choose_slot(){//слот который выбирает пациента, то есть ставит его на финальный слот
    if(ui->resultlabel->text()=="Success"){
        ui->pacientlabel->setText(ui->firstnameline->text());
        qDebug()<<"Success";
    }
    else{
        qDebug()<<"False";
    }
}

void Widget::addNewUser_slot(){//слот который добавляет нового пользователья в базу данных(в теории можно добавить всех пользователей в базу данных с помощью программы)
    open_if_open();
    if(ui->resultlabel->text()=="False"){
    query->exec("INSERT INTO Pacient VALUES ('"+ui->firstnameline->text().split(" ")[1]+"','"+ui->firstnameline->text().split(" ")[0]+"',123, 123);");
    }else{
        qDebug()<<"User is detected";
    }
}

void Widget::filling_slot(){//слот который заполняет распечатываемый документ exel и переводит его в формат pdf для последующего распечатываения
    QAxObject* printed = new QAxObject("Excel.Application", 0);
        QAxObject* printed_workbooks = printed->querySubObject("Workbooks");
        QString file = QDir::currentPath()+"\\printer_test.xlsx";
        file.replace("/","\\");
        QAxObject* printed_workbook = printed_workbooks->querySubObject("Open(const QString&)",  file);
        QAxObject* printed_sheets = printed_workbook->querySubObject("Worksheets");
        QAxObject* printed_sheet = printed_sheets->querySubObject("Item(const QVariant&)", 1);
        QAxObject* printed_cell;
        QStringList temp_list =ui->doclabel->text().split(" ");
        int j = 0;
        for(int i=3;i<8;i=i+2) {
            printed_cell = printed_sheet->querySubObject("Cells(int,int)", 17, i);
            printed_cell->setProperty("Value",temp_list[j]);
            ++j;
        }
        printed_cell = printed_sheet->querySubObject("Cells(int,int)", 15, 3);
        printed_cell->setProperty("Value",ui->pacientlabel->text());
        temp_list=QDate::currentDate().toString().split(" ");

        printed_cell = printed_sheet->querySubObject("Cells(int,int)", 13, 2);
        printed_cell->setProperty("Value",temp_list[2]);
        printed_cell = printed_sheet->querySubObject("Cells(int,int)", 13, 3);
        printed_cell->setProperty("Value",temp_list[1]);
        printed_cell = printed_sheet->querySubObject("Cells(int,int)", 13, 5);
        printed_cell->setProperty("Value",temp_list[3]);

        printed_cell = printed_sheet->querySubObject("Cells(int,int)", 19, 1);
        printed_cell->setProperty("Value",name_lec);

        qDebug()<<QDate::currentDate().toString();
        delete printed_cell;
        delete printed_sheet;
        qDebug()<<"izmeneno";
        delete printed_sheets;
        printed_workbook->dynamicCall("Save()");
        printed_workbook->dynamicCall("Close()");
        delete printed_workbook;
        //закрываем книги
        delete printed_workbooks;
        //закрываем Excel
        printed->dynamicCall("Quit()");
        delete printed;
        qDebug()<<"Finished";


}

void Widget::delete_user(){// слот который удаляет из базы данных человека, который в этой базе данных есть
    if(ui->resultlabel->text()=="Success"){
        query->exec("DELETE FROM Pacient WHERE Firstname = '"+ui->firstnameline->text().split(" ")[1]+"' AND Lastname = '"+ui->firstnameline->text().split(" ")[0]+"' ");
        qDebug()<< "User is deleted";
    }else{
        qDebug()<<"User is not detected";
    }

}

int Widget::sheets_name_f(QAxObject *ex,int sheetsCount){//функция которая находит номер листа с лекарством, если такого лекарства не существует, то оно возвразает -1
        for (int i=1;i<sheetsCount+1;i++)
        {
            QAxObject* sheetNames=ex->querySubObject("Item(const QVariant&)",QVariant(i));
            if(sheetNames->dynamicCall("Name").toString().toLower()==ui->lecarstvo->text().toLower()){
                qDebug()<<sheetNames->dynamicCall("Name").toString().toLower()<<" "<<i;
                qDebug()<<"Lecarstvo is detected";
                return i;
                break;
            }
            delete sheetNames;
        }
        return -1;
}

void Widget::set_db_lec(){//Запись выписанных лекарств в базу данных врачей

    db.close();
    db_lec = QSqlDatabase::addDatabase("QSQLITE");
    db_lec.setDatabaseName("./"+ui->doclabel->text().split(" ")[0]+".db");
    if(db_lec.open()){
        qDebug()<<"Open_doctors";
        query_lec = new QSqlQuery(db_lec);
        query_lec->exec("CREATE TABLE Шамиль("+ui->doclabel->text().split(" ")[0]+" TEXT;");
        qDebug()<<"CREATE TABLE "+ui->doclabel->text().split(" ")[0]+"("+ui->doclabel->text().split(" ")[0]+" TEXT;";
        query_lec->exec("INSERT INTO "+ui->doclabel->text().split(" ").at(0)+" VALUES ('"+name_lec+" для "+ui->pacientlabel->text()+" Дата "+QDate::currentDate().toString()+"');");
        qDebug()<<"INSERT INTO "+ui->doclabel->text().split(" ").at(0)+" VALUES ('"+name_lec+" для "+ui->pacientlabel->text()+"Дата"+QDate::currentDate().toString()+"');";
    }
    else {qDebug()<<"Not open";}

}

void Widget::set_lec(){//нахождение официального названия лекарства и заполнение третего окна приложения названием лекарства
        QAxObject* excel = new QAxObject("Excel.Application", 0);
        QAxObject* workbooks = excel->querySubObject("Workbooks");
        QString file = QDir::currentPath()+"\\test_4.xlsx";
        file.replace("/","\\");
        QAxObject* workbook = workbooks->querySubObject("Open(const QString&)",  file);
        QAxObject* sheets = workbook->querySubObject("Worksheets");

   int temp = sheets_name_f(sheets,sheets->property("Count").toInt());
   if(temp!=-1){
       qDebug()<<temp;
       QAxObject* sheet = sheets->querySubObject("Item(const QVariant&)", sheets_name_f(sheets,sheets->property("Count").toInt()));
       QAxObject* cell;
       cell = sheet->querySubObject("Cells(int,int)", 22, 1);
       QVariant temp= cell->property("Value");
       name_lec=temp.toString();

   delete cell;
   delete sheet;
   }
   delete sheets;
   delete workbook;
   //закрываем книги
   delete workbooks;
   //закрываем Excel
   excel->dynamicCall("Quit()");
   delete excel;
   if(name_lec=="") name_lec="Error_lecarstvo";
   qDebug()<<name_lec;
   ui->label_7->setText(name_lec);
}

void Widget::print(){//  планируется распечатка excel  документа
    QString fileName = "./1.exe";
    proc->start(fileName);
    qDebug()<<QDir::currentPath();
    proc->waitForFinished(30000);
}

void Widget::clear_slot(){// слот который очищает все поля программы кроме врача, потому что он сидит один за компьютером (наверное)
    ui->firstnameline->setText("");
    ui->resultlabel->setText("/*Статус*/");
    ui->podrobneelabel->setText("/*Полная Информация*/");
    ui->lecarstvo->setText("");
    ui->label_7->setText("");
    ui->pacientlabel->setText("");
}

void Widget::open_if_open(){//функция, которая проверяет ли открыта ли вторая дб, если да, то он ее закрывает и открывает первую
    if (db_lec.isOpen()){
        db_lec.close();
        db = QSqlDatabase::addDatabase("QSQLITE");
        db.setDatabaseName("./DB_for_amb.db");
        if(db.open()){
            qDebug()<<"Open";
        }
        else {qDebug()<<"Not open";}
    }
}

void Widget::perenos_slot(){//по факту бесполезная часть , так как скопировано для тестов
    QAxObject* printed = new QAxObject("Excel.Application", 0);
        QAxObject* printed_workbooks = printed->querySubObject("Workbooks");
        QAxObject* printed_workbook = printed_workbooks->querySubObject("Open(const QString&)",  "./printer_test.xlsx");
        QAxObject* printed_sheets = printed_workbook->querySubObject("Worksheets");

        QAxObject* printed_sheet = printed_sheets->querySubObject("Item(const QVariant&)", 1);
        QAxObject* printed_cell;
        QStringList temp_list =ui->doclabel->text().split(" ");
        int j = 0;
        for(int i=3;i<8;i=i+2) {
            printed_cell = printed_sheet->querySubObject("Cells(int,int)", 17, i);
            printed_cell->setProperty("Value",temp_list[j]);
            ++j;
        }
        printed_cell = printed_sheet->querySubObject("Cells(int,int)", 15, 3);
        printed_cell->setProperty("Value",ui->pacientlabel->text());
        temp_list=QDate::currentDate().toString().split(" ");

        printed_cell = printed_sheet->querySubObject("Cells(int,int)", 13, 2);
        printed_cell->setProperty("Value",temp_list[2]);
        printed_cell = printed_sheet->querySubObject("Cells(int,int)", 13, 3);
        printed_cell->setProperty("Value",temp_list[1]);
        printed_cell = printed_sheet->querySubObject("Cells(int,int)", 13, 5);
        printed_cell->setProperty("Value",temp_list[3]);

        printed_cell = printed_sheet->querySubObject("Cells(int,int)", 19, 1);
        printed_cell->setProperty("Value",name_lec);
        //нижний dynamicCall форматирует из exel в pdf для печати
        printed_sheet->dynamicCall("ExportAsFixedFormat(int, const QString&, int, BOOL, BOOL)", 0, "./example.pdf", 0, false, false);
        //

        //
        qDebug()<<QDate::currentDate().toString();
        delete printed_cell;
        delete printed_sheet;

        qDebug()<<"izmeneno";

    delete printed_sheets;
    printed_workbook->dynamicCall("Save()");
    delete printed_workbook;
    //закрываем книги
    delete printed_workbooks;
    //закрываем Excel
    printed->dynamicCall("Quit()");
    delete printed;
    qDebug()<<"Finished";


}
Widget::~Widget()
{
    delete ui;
    delete query;
    if(query_lec!=NULL) delete query_lec;
}
