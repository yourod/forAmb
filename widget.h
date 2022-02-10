#pragma once
#include <QWidget>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QDebug>
#include <QSqlTableModel>
#include <QAxObject>
#include <QStringList>
#include <QTableWidget>
#include <QProcess>
QT_BEGIN_NAMESPACE
namespace Ui { class Widget; }
QT_END_NAMESPACE

class Widget : public QWidget
{
    Q_OBJECT

public:
    Widget(QWidget *parent = nullptr);
    ~Widget();

private slots:

    void FIO_slot();
    void search_slot();
    void choose_slot();
    void addNewUser_slot();
    void filling_slot();
    void delete_user();
    void set_db_lec();
    void set_lec();
    void print();
    void clear_slot();
    void perenos_slot();
private:
    Ui::Widget *ui;
    QSqlDatabase db;
    QSqlQuery* query;
    QSqlDatabase db_lec;
    QSqlQuery* query_lec;
    int row;
    QStringList L;
    QMap<QString, int> example;
    QString name_lec;
    int sheets_name_f(QAxObject* ex,int sheetsCount);
    QProcess *proc;
    void open_if_open();

};
