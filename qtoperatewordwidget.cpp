#include "qtoperatewordwidget.h"
#include "ui_qtoperatewordwidget.h"
#include<QFileDialog>
#include <QDir>
#include<QDebug>
#include <QMessageBox>
#include <QDateTime>
QtOperateWordWidget::QtOperateWordWidget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::QtOperateWordWidget)
{
    ui->setupUi(this);
    m_pWord=nullptr;
    m_pDocuments=nullptr;
    m_pDocument=nullptr;


//    QStringList headList;
//    headList<<"head1111111"<<"head2"<<"head3"<<"head4"<<"head5";
//    QAxObject*table=CreateTable(5, 5,  headList);
QString path1="C:/Users/Administrator.SC-201905281555/Desktop/66677.docx";
QFile  file(path1);
//    CreateWordDocument(path1);

QAxObject*words=new QAxObject(this);

words->setControl("Word.Application");
//m_pWord=new QAxWidget("Word.Application");//新建一个word应用程序
words->setProperty("Visible",  false);//不显示窗口
QAxObject *pDocuments=words->querySubObject("Documents");
//    pDocuments->dynamicCall("Add(QString)",QString("C:/Users/Administrator.SC-201905281555/Desktop/test.docx"));//获取对应目录一个模板
//pDocuments->dynamicCall("Add(QString)",path1);//获取对应目录一个模板
pDocuments->dynamicCall("Add(void)");//获取对应目录一个模板
m_pDocument=words->querySubObject("ActiveDocument");//获取当前激活文档





//    // QString  h="测试";
//    QString  a="姓名";
//    TagListMatching.append(a);
//    QString  b="性别";
//    TagListMatching.append(b);
//    QString  c="年龄";
//    TagListMatching.append(c);
//    QString   d="民族";
//    TagListMatching.append(d);
//    QString   e="出生年月";
//    TagListMatching.append(e);
//    QString   f="学历";
//    TagListMatching.append(f);
//    DataTextList.append("xao");
//    DataTextList.append("女");
//    DataTextList.append("23");
//    DataTextList.append("汉");
//    DataTextList.append("199/2/3");
//    DataTextList.append("本科");
//    QString  content="123;456;234;123;99;22";
//    //    InsertText(TagListMatching, DataTextList);
//    //    //    QString path="D:\\123.doc";

//    WordTagDatas.insert("姓名","小明");
//    WordTagDatas.insert("年龄","18");
//    WordTagDatas.insert("民族","han");
//    WordTagDatas.insert("出生年月","199/2/3");
//    WordTagDatas.insert("出生年月","199/2/1");
//    WordTagDatas.insert("学历","本科");
//    WordTagDatas.insert("学历","本科");
//    WordTagDatas.insert("民族","an");

//    LastInsertText(WordTagDatas);

//    QStringList headList;
//    headList<<"head1111111"<<"head2"<<"head3"<<"head4"<<"head5";
//    QAxObject*table=CreateTable(5, 5,  headList);

// InsertTable(0,0, 5, 8);
//    QString path="D:\\123.docx";
      QString path="C:\\Users\\Administrator.SC-201905281555\\Desktop\\147.docx";
//    SaveAndQuit(path1);

//      m_pDocument->dynamicCall("SaveAs(const QString &text)",QDir::toNativeSeparators(path1));//保存文档
         m_pDocument->dynamicCall("SaveAs(const QString &)",QDir::toNativeSeparators(path1));//保存文档
      //    m_pDocument->dynamicCall("Close(boolean)",true);//关闭
      m_pDocument->dynamicCall("Close(bool)",true);//关闭
      //    m_pDocument->dynamicCall("Quit()");
      words->dynamicCall("Quit()");
file.close();
}

QtOperateWordWidget::~QtOperateWordWidget()
{
    delete ui;
}
///创建一个文档
void QtOperateWordWidget::CreateWordDocument( QString path1)
{


    m_pWord=new QAxWidget("Word.Application");//新建一个word应用程序
    m_pWord->setProperty("Visible",  false);//不显示窗口
    QAxObject *pDocuments=m_pWord->querySubObject("Documents");
//    pDocuments->dynamicCall("Add(QString)",QString("C:/Users/Administrator.SC-201905281555/Desktop/test.docx"));//获取对应目录一个模板
    pDocuments->dynamicCall("Add(QString)",path1);//获取对应目录一个模板
    m_pDocument=m_pWord->querySubObject("ActiveDocument");//获取当前激活文档

}

bool QtOperateWordWidget::InsertText( QList<QString>TagList, QList<QString>Text)
{
    if(m_pDocument->isNull())//首先判断有没有当前激活文档，没有返回失败
        return false;
    for (int i=0;i<TagList.size();i++) {
        QAxObject *pBookMarkCode=m_pDocument->querySubObject("Bookmarks(QVariant)",TagList.at(i));//查询获取指定标签
        //        for (int k=0;k<Text.size();k++) {

        if(pBookMarkCode){
            pBookMarkCode->dynamicCall("Select(void)");//选择该指定标签
            pBookMarkCode->querySubObject("Range")->setProperty("Text",Text.at(i));//往标签处插入文字
            delete pBookMarkCode;
        }

        //    QAxObject *pBookMarkCode=m_pDocument->querySubObject("Bookmarks(QVariant)",Tag);//查询获取指定标签
        //   QAxObject *pBookMarkCode=m_pDocument->querySubObject("Bookmarks(QString)",Tag);//获取指定标签
        //    if(pBookMarkCode){
        //        pBookMarkCode->dynamicCall("Select(void)");//选择该指定标签
        //        pBookMarkCode->querySubObject("Range")->setProperty("Text",text);//往标签处插入文字
        //        delete pBookMarkCode;
        //      return true;

        //        }
        //    return true;
    }
    //    return false;
    return true;

}

void QtOperateWordWidget::SaveAndQuit(const QString &text)
{
    m_pDocument->dynamicCall("SaveAs(const QString &text)",QDir::toNativeSeparators(text));//保存文档
    //    m_pDocument->dynamicCall("Close(boolean)",true);//关闭
    m_pDocument->dynamicCall("Close(bool)",true);//关闭
    //    m_pDocument->dynamicCall("Quit()");
    m_pWord->dynamicCall("Quit()");
}
///通过标签和数据键值对向WORD中加入数据
bool QtOperateWordWidget::LastInsertText(QMap<QString, QString> MatchMap)
{
    if(m_pDocument->isNull())//首先判断有没有当前激活文档，没有返回失败
        return false;
    //    MatchMap.begin();
    QMap<QString, QString> ::const_iterator i=MatchMap.constBegin();
    while (i!=MatchMap.constEnd()) {
        QAxObject *pBookMarkCode=m_pDocument->querySubObject("Bookmarks(QVariant)",i.key());//查询获取指定标签
        if(pBookMarkCode){
            pBookMarkCode->dynamicCall("Select(void)");//选择该指定标签
            pBookMarkCode->querySubObject("Range")->setProperty("Text",i.value());//往标签处插入文字
            delete pBookMarkCode;
            i++;
        }

    }
    return  true;

}
///生成动态的word表格并填入数据
void QtOperateWordWidget::CreateDynamicWordDocument()
{






}
/////创建表格
QAxObject* QtOperateWordWidget::CreateTable(int row, int column, QStringList headList)
{
       QAxObject *selection=m_pWord->querySubObject("Selection");

       if(!selection)
          { return nullptr;
       }
       selection->dynamicCall("InsertAfter(QString&)","\r\n");
       QAxObject*range=selection->querySubObject("Range");
       QAxObject*tables=m_pDocument->querySubObject("Tables");
       QAxObject*table= tables->querySubObject("Add(QVariant,int ,int)",range->asVariant(),row,column);
      table->setProperty("Style","网格型");
      table->dynamicCall("AutoFitBehavior(WdAutoFitBehavior)",2);
      for (int i=0;i<headList.size();i++) {
          table->querySubObject("Cell(int,int)",1,i+1)->querySubObject("Range")->dynamicCall("setText(Qstring)",headList.at(i));
         table->querySubObject("Cell(int,int)",1,i+1)->querySubObject("Range")->dynamicCall("setBold(int)",true);

      }


      return table;

}
/////创建表格
//QAxWidget *QtOperateWordWidget::CreateTable(int row, int column, QStringList headList)
//{

////     QAxObject *Selection=m_pDocument->querySubObject("Range(Long ,Long)",nStart,nEnd);



//}


///向word中插入表格
void QtOperateWordWidget::InsertTable(int nStart, int nEnd, int row, int column)
{
    QAxObject *ptst=m_pDocument->querySubObject("Range(Long ,Long)",nStart,nEnd);
//  QAxObject *selection=m_pWord->querySubObject("Selection");
    QAxObject *pTable=m_pDocument->querySubObject("Tables");
    QVariantList params;
    params.append(ptst->asVariant());
    params.append(row);
    params.append(column);
    if(pTable)
    {
       pTable->dynamicCall("Add(QAxObject*,Long ,Long)",params);
  QAxObject*table= pTable->querySubObject("Add(QAxObject*,Long ,Long)",params);
//       QAxObject*table= selection->querySubObject("Tables(int),1)",params);
//       table->dynamicCall("Style","网格型");
//       table->dynamicCall("AutoFitBehavior(WdAutoFitBehavior)",2);


       QAxObject* Borders=table->querySubObject("Borders");
       Borders->setProperty("InsideLineStyle",1);
       Borders->setProperty("OutsideLineStyle",1);

    }

//pTable->dynamicCall("Style","网格型");
}

///采用CSV文件格式保存文件
void QtOperateWordWidget::on_pushButton_clicked()
{
    //提示用户导出，获取导出路径
    QString sFileName=QFileDialog::getSaveFileName(this,"选择保存路径","","csv 文件(*.csv)");
    if(!sFileName.isNull()){
        QFile *pFile=new QFile(sFileName);
        QDateTime dcurentTime=QDateTime::currentDateTime();
        QString  h=dcurentTime.toString("yyyy.MM.dd hh:mm:ss");
        //  QString sTmp=QString("界面%1 ,项目:%2,测试结果:%3,测试时间:%4").arg("低频插件").arg("直流电源").arg("成功").arg(h);
        QString sTmp=QString("界面 ,项目,测试结果,测试时间\n");
        QString sTmp1 =QString("%1 ,%2,%3,%4\n").arg("低频插件").arg("直流电源").arg("成功").arg(h);
        qDebug()<<sTmp;
        CSVDataList.append(sTmp);
        CSVDataList.append(sTmp1);

        if(pFile->open(QIODevice::WriteOnly)){
            QString sFiles=sTmp;
            for (int i=0;i<CSVDataList.size();i++) {
                pFile->write(CSVDataList.at(i).toLocal8Bit().data());
            }
            //              pFile->write(sFiles.toLocal8Bit().data());
            CSVDataList.clear();
            pFile->close();
            QMessageBox::information(this,"提示","数据导出成功");
        }
    }


}
