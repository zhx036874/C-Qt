#ifndef QTOPERATEWORDWIDGET_H
#define QTOPERATEWORDWIDGET_H

#include <QWidget>
#include<QAxObject>
#include<QAxWidget>
namespace Ui {
class QtOperateWordWidget;
}

class QtOperateWordWidget : public QWidget
{
    Q_OBJECT

public:
    explicit QtOperateWordWidget(QWidget *parent = nullptr);
    ~QtOperateWordWidget();

public:
    void CreateWordDocument(QString path1);//创建word
    bool InsertText(QList<QString>TagList,QList<QString> Text);//向标签处插入文字
    void SaveAndQuit(const QString &text);//保存并退出
    ///向WORD中加入数据
 bool LastInsertText( QMap<QString ,QString>MatchMap);//向标签处插入文字
    ///导入对应wordtag的数据
QList<QString>DataTextList;
    ///设置标签列表
  QList<QString>TagListMatching;
    ///利用CSV处理测试报告定义开始
   QList<QString>CSVDataList;
    ///
 ///利用CSV处理测试报告定义结束
 ///定义数据暂存键值对
 QMap<QString ,QString>WordTagDatas;


 ///生成动态的word表格并填入数据
 void  CreateDynamicWordDocument();
///创建表格
 QAxObject* CreateTable( int row ,int column,QStringList headList  );
 ///向word中插入表格
 void InsertTable(int nStart, int nEnd ,int row ,int column);
private slots:
    void on_pushButton_clicked();

private:
    QString m_fileName;//存入位置
    QAxWidget*m_pWord;
    QAxObject*m_pDocuments;
     QAxObject*m_pDocument;
    Ui::QtOperateWordWidget *ui;
};


#if _MSC_VER>=1600 //中文字符宏定义，避免乱码出现
#pragma execution_character_set("utf-8")
#endif
#endif // QTOPERATEWORDWIDGET_H
