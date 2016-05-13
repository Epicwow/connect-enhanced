#include "connect.h"
#include "ui_connect.h"
#include <QMenu>
#include <QValidator>
#include <QStringList>
#include <QCompleter>
#include <QList>
#include <QDateTime>
#include "xlsxdocument.h"
#include <QFileDialog>
#include <QInputDialog>
#include <QTcpSocket>

#define IP_TAIL_MAX 50

Connect::Connect(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::Connect)
{
    // 默认变量
    server_ip = "192.168.1.5";
    ui->setupUi(this);
    status_label = new QLabel;
    if_query_is_set = 0;
    if_connected = false;
    dbid = 0;
    hide_search_table();
    ui->fgroupWidget->hide();
    ui->lineEditRace->setText("汉族");
    ui->lineEditProvince->setText("北京市");
    ui->lineEditCity->setText("北京市");
    ui->lineEditQuery->hide();
    //ui->labelEditor->hide();

    //completor
    {

        // 性别
        QStringList gender_list;
        gender_list << " 1. 男" << " 2. 女";
        QCompleter *gender_completer = new QCompleter(gender_list, this);
        ui->lineEditGender->setCompleter(gender_completer);

        // 民族
        QStringList race_list;
        race_list << " acz 阿昌族" << " bz 白族" << " baz 保安族" << " blz 布朗族" << " cxz 朝鲜族" << " dwez 达斡尔族" << " dz 傣族" << " daz 德昂族"
                  << " dxz 东乡族" << " dz 侗族" << " dlz 独龙族" << " elsz 俄罗斯族" << " elcz 鄂伦春族" << " ewkz 鄂温克族" << " gsz 高山族" << " glz 仡佬族"
                  << " hnz 哈尼族" << " hskz 哈萨克族" << " hz 汉族" << " hzz 赫哲族" << " hz 回族" << " jnz 基诺族" << " jz 京族" << " jpz 景颇族" << " hekzz 哈尔克孜族"
                  << " lhz 拉祜族" << " lz 黎族" << " lsz 傈僳族" << " lbz 珞巴族" << " mz 满族" << " mnz 毛南族" << " mbz 门巴族" << " mgz 蒙古族" << " mz 苗族"
                  << " mlz 仫佬族" << " nxz 纳西族" << " nz 怒族" << " pmz 普米族" << " qz 羌族" << " slz 撒拉族" << " sz 畲族" << " sz 水族" << " tjkz 塔吉克族"
                  << " ttez 塔塔尔族" << "tjz 土家族" << " tz 土族" << " wz 佤族" << " wwez 维吾尔族" << " wzbkz 乌孜别克族" << " xbz 锡伯族" << " yz 瑶族"
                  << " yz 彝族" << " ygz 裕固族" << " zz 藏族" << " zz 壮族";
        QCompleter *race_completer = new QCompleter(race_list, this);
        race_completer->setCaseSensitivity(Qt::CaseInsensitive);
        ui->lineEditRace->setCompleter(race_completer);

        // 文化程度
        QStringList degree_list;
        degree_list << "小学" << "初中" << "高中" << "高中中专" << "大学专科" << "大学本科" << "硕士" << "博士" << "博士后";
        QCompleter *degree_completer = new QCompleter(degree_list, this);
        degree_completer->setCompletionMode(QCompleter::UnfilteredPopupCompletion);
        ui->lineEditDegree->setCompleter(degree_completer);

        // 健康状况
        QStringList health_list;
        health_list << " 1. 良好" << " 2. 一般" << " 3. 欠佳" << " 4. 重病";
        QCompleter *health_completer = new QCompleter(health_list, this);
        health_completer->setCompletionMode(QCompleter::UnfilteredPopupCompletion);
        ui->lineEditHealth->setCompleter(health_completer);

        // 是否深刻了解佛法
        QStringList knowledge_list;
        knowledge_list << " 1. 完全不了解" << " 2. 不深刻" << " 3. 比较深刻";
        QCompleter *knowledge_completer = new QCompleter(knowledge_list, this);
        knowledge_completer->setCompletionMode(QCompleter::UnfilteredPopupCompletion);
        ui->lineEditBuddhismLevel->setCompleter(knowledge_completer);

        // 以何因缘了解佛法
        QStringList reason_list;
        reason_list << " 1. 由法师(居士)引导学佛" << " 2. 由读佛教经典领悟学佛" << " 3. 由经历启发学佛";
        QCompleter *reason_completer = new QCompleter(reason_list, this);
        reason_completer->setCompletionMode(QCompleter::UnfilteredPopupCompletion);
        ui->lineEditReasonToLearnDharma->setCompleter(reason_completer);

        // 学佛小组种类
        QStringList learn_kinds_list;
        learn_kinds_list << " 1. 周末山上 " << " 2. 平常山下" << " 3. 京外" << " 4. 境外" << " 5. 外语";
        QCompleter *learn_kinds_completer = new QCompleter(learn_kinds_list, this);
        learn_kinds_completer->setCompletionMode(QCompleter::UnfilteredPopupCompletion);
        ui->lineEditLearnKind->setCompleter(learn_kinds_completer);
    }

    // left table view
    {
        viewModel = new QStandardItemModel();
        viewModel->setHorizontalHeaderItem(0, new QStandardItem(QObject::trUtf8("姓名")));
        viewModel->setHorizontalHeaderItem(1, new QStandardItem(QObject::trUtf8("手机")));
        viewModel->setHorizontalHeaderItem(2, new QStandardItem(QObject::trUtf8("收据编号")));
        viewModel->setHorizontalHeaderItem(3, new QStandardItem(QObject::trUtf8("皈依证号")));
        viewModel->setHorizontalHeaderItem(4, new QStandardItem(QObject::trUtf8("身份证号码")));
        viewModel->setHorizontalHeaderItem(5, new QStandardItem(QObject::trUtf8("固定电话")));
        viewModel->setHorizontalHeaderItem(6, new QStandardItem(QObject::trUtf8("省")));
        viewModel->setHorizontalHeaderItem(7, new QStandardItem(QObject::trUtf8("市")));
        viewModel->setHorizontalHeaderItem(8, new QStandardItem(QObject::trUtf8("区/县")));
        viewModel->setHorizontalHeaderItem(9, new QStandardItem(QObject::trUtf8("街道")));
        ui->tableView->setModel(viewModel);
        ui->tableView->setEditTriggers(QAbstractItemView::NoEditTriggers);
    }


    {
        viewModel_search = new QStandardItemModel();
        viewModel_search->setHorizontalHeaderItem(0, new QStandardItem(QObject::trUtf8("姓名")));
        viewModel_search->setHorizontalHeaderItem(1, new QStandardItem(QObject::trUtf8("手机")));
        viewModel_search ->setHorizontalHeaderItem(2, new QStandardItem(QObject::trUtf8("收据编号")));
        viewModel_search->setHorizontalHeaderItem(3, new QStandardItem(QObject::trUtf8("皈依证号")));
        viewModel_search->setHorizontalHeaderItem(4, new QStandardItem(QObject::trUtf8("身份证号码")));
        viewModel_search->setHorizontalHeaderItem(5, new QStandardItem(QObject::trUtf8("固定电话")));
        viewModel_search->setHorizontalHeaderItem(6, new QStandardItem(QObject::trUtf8("省")));
        viewModel_search->setHorizontalHeaderItem(7, new QStandardItem(QObject::trUtf8("市")));
        viewModel_search->setHorizontalHeaderItem(8, new QStandardItem(QObject::trUtf8("区/县")));
        viewModel_search->setHorizontalHeaderItem(9, new QStandardItem(QObject::trUtf8("街道")));


        ui->tableView_2->setModel(viewModel_search);
        ui->tableView_2->setEditTriggers(QAbstractItemView::NoEditTriggers);
    }

    // Regex to test if insert is right
    {
        /* set validators [below] */

        QRegExp regExpPhoneNum("1[3|4|5|7|8|][0-9]{9}"); // 判断输入是否为手机号，只允许输入正常手机号
        ui->lineEditPhoneNum->setValidator(new QRegExpValidator(regExpPhoneNum, this));
        ui->lineEdit_Phone->setValidator(new QRegExpValidator(regExpPhoneNum, this));

        /*
        QRegExp regExpTel("\\d{3}-\\d{8}|\\d{4}-\\d{7}");
        ui->lineEditTelephoneNum->setValidator(new QRegExpValidator(regExpTel, this));
        */

        QRegExp regExpBirthday("^(19|20)\\d{2}.(1[0-2]|0?[1-9]).(0?[1-9]|[1-2][0-9]|3[0-1])$");
        ui->lineEditBirthday->setValidator(new QRegExpValidator(regExpBirthday, this));
        ui->lineEdit_Birthday->setValidator(new QRegExpValidator(regExpBirthday, this));

        QRegExp regExpEditTime("^(19|20)\\d{2}.(1[0-2]|0?[1-9]).(0?[1-9]|[1-2][0-9]|3[0-1])$");
        ui->lineEditTime->setValidator(new QRegExpValidator(regExpEditTime, this));

        QRegExp regExpGraduateTime("^(19|20)\\d{2}.(1[0-2]|0?[1-9]).(0?[1-9]|[1-2][0-9]|3[0-1])$");
        ui->lineEditGraduateTime->setValidator(new QRegExpValidator(regExpGraduateTime, this));

        QRegExp regExpFirstJobDate("^(19|20)\\d{2}.(1[0-2]|0?[1-9]).(0?[1-9]|[1-2][0-9]|3[0-1])$");
        ui->lineEditFirstJobDate->setValidator(new QRegExpValidator(regExpFirstJobDate, this));

        QRegExp regExpSecondJobDate("^(19|20)\\d{2}.(1[0-2]|0?[1-9]).(0?[1-9]|[1-2][0-9]|3[0-1])$");
        ui->lineEditSecondJobDate->setValidator(new QRegExpValidator(regExpSecondJobDate, this));

        QRegExp regExpRetirementDate("^(19|20)\\d{2}.(1[0-2]|0?[1-9]).(0?[1-9]|[1-2][0-9]|3[0-1])$");
        ui->lineEditRetirementDate->setValidator(new QRegExpValidator(regExpRetirementDate, this));

        /* 去掉对身份证的校验
        QRegExp regExpID("^[1-9]\\d{5}[1-9]\\d{3}((0\\d)|(1[0-2]))(([0|1|2]\\d)|3[0-1])((\\d{4})|\\d{3}[Xx])$");
        ui->lineEditID->setValidator(new QRegExpValidator(regExpID, this));
        */

        QRegExp regExpYearsOfLearningBuhddhism("^[1-9]\\d?$");
        ui->lineEditYearsOfLearningBuddhism->setValidator(new QRegExpValidator(regExpYearsOfLearningBuhddhism, this));

        /* 去掉对邮编的校验
        QRegExp regExpPostcode("^[1-9][0-9]{5}$");
        ui->lineEditPostcode->setValidator(new QRegExpValidator(regExpPostcode, this));
        */
        /* set validators [above] */
    }

    /* set edit time [below] */
    QDateTime time = QDateTime::currentDateTime();
    QString edit_date = time.toString("yyyy.MM.dd");
    ui->lineEditTime->setText(edit_date);
    /* set edit time [above] */

    // set local ip at status bar
    get_local_ip();

    if (local_ip.section(".", -1).toInt() > IP_TAIL_MAX) {
        ui->actionInitDb->setDisabled(true);
    }

    status_label->setStyleSheet("font-size:9px");
    status_label->setText(QString(" Local Address: [%1], Server Default Address: [%2]").arg(local_ip, server_ip));
    statusBar()->addWidget(status_label);
}

Connect::~Connect()
{
    delete ui;
}

bool Connect::modify_or_not()
{
    QMessageBox::StandardButton ask
            = QMessageBox::question(this,
                                    "connect",
                                    "你确定更改此行内容么? 如果更改请按 Yes，否则按 Cancel，更改后请保存\n",
                                    QMessageBox::Cancel|QMessageBox::Yes);
    if (ask == QMessageBox::Yes)
        return true;
    else
        return false;
}

void Connect::on_tableView_doubleClicked(const QModelIndex &index)
{
    int order;
    if (!ui->actionJoinin->isEnabled()) {
        order = 0;
    } else {
        order = 2; // 第二列是 receipt
    }

    QString receipt = index.sibling(index.row(), order).data().toString();

    if (modify_or_not()) {
        complete_fields("receipt", receipt);
        viewModel->removeRow(index.row());
    } else return;
}

void Connect::on_lineEditReceipt_editingFinished()
{
    if (!test_if_connected()) return;
    if (ui->lineEditReceipt->text().isEmpty()) return;
    complete_fields("receipt", ui->lineEditReceipt->text().toUpper());
}

bool Connect::check_lineEdit_items()
{
    bool name;
    bool editor;

    if (!ui->actionJoinin->isEnabled()) {
        name = ui->lineEdit_Name->text().isEmpty();
        editor = ui->lineEditOthers->text().isEmpty();
        qDebug() << "check items";
    } else {
        name = ui->lineEditName->text().isEmpty();
        editor = ui->lineEditor->text().isEmpty();
    }

    if (name || editor) return false;
    else return true;
}

void Connect::on_pushButtonOK_clicked()
{
    if (!test_if_connected()) {
        qDebug() << "database is not connected! please connect again!";
        return;
    }

    if (!check_lineEdit_items()) {
        QMessageBox::information(this, "", "请至少填写姓名、编辑人");
        return;
    }

    validate_input_values();
    if (update_database()) {
        append_items2_tableView();
        clear_lineEdits();
    }

    if (!ui->actionJoinin->isEnabled()) {
        ui->lineEdit_Name->setFocus();
    } else {
        ui->lineEditName->setFocus();
    }
}

/* user defined functions */
bool Connect::validate_input_values()
{
    //[1] valid ID card
    if (ui->lineEditID->text().isEmpty()) {
        qDebug() << "ID empty";
    } else {
        if (ui->lineEditID->text().length() != 18) {
            QMessageBox::information(this, "", "身份证号必须是18位得组合");
            /* 如果身份证输入错误，提交的时候会弹出对话框，确定后会把焦点转移到身份证输入框 */
            ui->lineEditID->setFocus();
        }
    }

    //[2] valid phone num 手机号必须位数为11位数字，正则无法检测输入比这个短
    if (ui->lineEditPhoneNum->text().isEmpty()) {
        qDebug() << "phone num is empty";
    } else {
        if (ui->lineEditPhoneNum->text().length() != 11) {
            QMessageBox::information(this, "", "手机号必须是11位");
            ui->lineEditPhoneNum->setFocus();
        }
    }

    return true;
}

bool Connect::update_database()
{
    qDebug() << "xxx dbid = " << dbid;
    QSqlQuery query;

    if (!ui->actionJoinin->isEnabled()) {
        if (dbid != 0) {
            query.prepare("UPDATE zen_person set receipt = :receipt,"
                          "name = :name,"
                          "gender = :gender,"
                          "birthday = :birthday,"
                          "phone_num = :phone_num,"
                          "telephone_num = :telephone_num,"
                          "learn_dharma_kinds = :learn_dharma_kinds,"
                          "province = :province,"
                          "city = :city,"
                          "district = :district,"
                          "address = :address,"
                          "if_apply_learn_place = :if_apply_learn_place,"
                          "notes = :notes,"
                          "others = :others, "
                          "learn_dharma_address = :learn_dharma_address "
                          "where id = :dbid"
                          );
        } else {
            query.prepare("INSERT into zen_person ("
                          "receipt,"
                          "name,"
                          "gender,"
                          "birthday,"
                          "phone_num,"
                          "telephone_num,"
                          "learn_dharma_kinds,"
                          "province,"
                          "city,"
                          "district,"
                          "address,"
                          "if_apply_learn_place,"
                          "notes,"
                          "others,"
                          "learn_dharma_address"
                          ")"
                          "values ("
                          ":receipt,"
                          ":name,"
                          ":gender,"
                          ":birthday,"
                          ":phone_num,"
                          ":telephone_num,"
                          ":learn_dharma_kinds,"
                          ":province,"
                          ":city,"
                          ":district,"
                          ":address,"
                          ":if_apply_learn_place,"
                          ":notes,"
                          ":others,"
                          ":learn_dharma_address)"
                          );
        }

        query.bindValue(":receipt", ui->lineEdit_Order->text());
        query.bindValue(":name", ui->lineEdit_Name->text());
        query.bindValue(":gender", ui->lineEdit_Gender->text());
        query.bindValue(":birthday", ui->lineEdit_Birthday->text());
        query.bindValue(":phone_num", ui->lineEdit_Phone->text());
        query.bindValue(":telephone_num", ui->lineEdit_ContractWay->text());
        query.bindValue(":learn_dharma_kinds", ui->lineEditLearnKind->text());
        query.bindValue(":province", ui->lineEdit_Province->text());
        query.bindValue(":city", ui->lineEdit_District->text());
        query.bindValue(":district", ui->lineEdit_District->text());
        query.bindValue(":address", ui->lineEdit_Address->text());
        query.bindValue(":if_apply_learn_place", ui->lineEdit_ApplyPlace->text());
        query.bindValue(":notes", ui->lineEdit_Note->text());
        query.bindValue(":others", ui->lineEditOthers->text());
        query.bindValue(":learn_dharma_address", ui->lineEditLearnAddress->text());
        query.bindValue(":dbid", QString("%1").arg(dbid));

        if (!query.exec()) {
            QString error_msg = query.lastError().text();
            QMessageBox::information(this, "", QString("数据库错误，请联系开发运维人员：") + error_msg);
            qDebug() << query.executedQuery();
            return false;
        }

        query.clear();
        return true;
    }

    if (dbid != 0) {
        query.prepare("\
                      UPDATE `zen_person` \
                      SET \
                      `name` = :name ,\
                      `gender` = :gender ,\
                      `job` = :job ,\
                      `hobby` = :hobby ,\
                      `fname` = :fname ,\
                      `birthday` = :birthday ,\
                      `personnel_id` = :personnel_id ,\
                      `phone_num` = :phone_num ,\
                      `race` = :race ,\
                      `degree` = :degree ,\
                      `health` = :health ,\
                      `telephone_num` = :telephone_num ,\
                      `edit_time` = :edit_time ,\
                      `receipt` = :receipt ,\
                      `workplace` = :workplace ,\
                      `province` = :province ,\
                      `city` = :city ,\
                      `district` = :district ,\
                      `address` = :address ,\
                      `postcode` = :postcode ,\
                      `graduate_time` = :graduate_time ,\
                      `graduate_school` = :graduate_school ,\
                      `first_job_entry_time` = :first_job_entry_time ,\
                      `first_job_workplace` = :first_job_workplace ,\
                      `second_job_entry_time` = :second_job_entry_time ,\
                      `second_job_workplace` = :second_job_workplace ,\
                      `retirement_date` = :retirement_date ,\
                      `retirement_workplace` = :retirement_workplace ,\
                      `year2start_learning_buddhism` = :year2start_learning_buddhism ,\
                      `years_of_learning_buddhism` = :years_of_learning_buddhism ,\
                      `deep_understanding_of_dharma` = :deep_understanding_of_dharma ,\
                      `reason2learning_dharma` = :reason2learning_dharma ,\
                      `nums_of_buddhism_book` = :nums_of_buddhism_book ,\
                      `easy2learn_buddhism_book` = :easy2learn_buddhism_book ,\
                      `hard2read` = :hard2read ,\
                      `maxim` = :maxim ,\
                      `buddhist_disciples_of_family` = :buddhist_disciples_of_family ,\
                      `editor` = :editor ,\
                      `code` = :code \
                WHERE `id` = :dbid; "
                );
    } else {
        query.prepare(
                    "INSERT INTO `zen_person`\
                    (   `name`,\
                        `gender`,\
                        `job`,\
                        `hobby`,\
                        `fname`,\
                        `birthday`,\
                        `personnel_id`,\
                        `phone_num`,\
                        `race`,\
                        `degree`,\
                        `health`,\
                        `telephone_num`,\
                        `edit_time`,\
                        `receipt`,\
                        `workplace`,\
                        `province`,\
                        `city`,\
                        `district`,\
                        `address`,\
                        `postcode`,\
                        `graduate_time`,\
                        `graduate_school`,\
                        `first_job_entry_time`,\
                        `first_job_workplace`,\
                        `second_job_entry_time`,\
                        `second_job_workplace`,\
                        `retirement_date`,\
                        `retirement_workplace`,\
                        `year2start_learning_buddhism`,\
                        `years_of_learning_buddhism`,\
                        `deep_understanding_of_dharma`,\
                        `reason2learning_dharma`,\
                        `nums_of_buddhism_book`,\
                        `easy2learn_buddhism_book`,\
                        `hard2read`,\
                        `maxim`,\
                        `buddhist_disciples_of_family`,\
                        `editor`,\
                        `code`,\
                         `ipaddress`\
                              )\
                                \
                    VALUES\
                    (   :name,\
                        :gender,\
                        :job,\
                        :hobby,\
                        :fname,\
                        :birthday,\
                        :personnel_id,\
                        :phone_num,\
                        :race,\
                        :degree,\
                        :health,\
                        :telephone_num,\
                        :edit_time,\
                        :receipt,\
                        :workplace,\
                        :province,\
                        :city,\
                        :district,\
                        :address,\
                        :postcode,\
                        :graduate_time,\
                        :graduate_school,\
                        :first_job_entry_time,\
                        :first_job_workplace,\
                        :second_job_entry_time,\
                        :second_job_workplace,\
                        :retirement_date,\
                        :retirement_workplace,\
                        :year2start_learning_buddhism,\
                        :years_of_learning_buddhism,\
                        :deep_understanding_of_dharma,\
                        :reason2learning_dharma,\
                        :nums_of_buddhism_book,\
                        :easy2learn_buddhism_book,\
                        :hard2read,\
                        :maxim,\
                        :buddhist_disciples_of_family,\
                        :editor,\
                        :code,\
                        :ipaddress\
                        );"
                );
    }

    query.bindValue(":name", ui->lineEditName->text());
    query.bindValue(":gender", ui->lineEditGender->text().section(' ', -1));
    query.bindValue(":job", ui->lineEditJob->text());
    query.bindValue(":hobby", ui->lineEditHobby->text());
    query.bindValue(":fname", ui->lineEditFName->text());
    query.bindValue(":birthday", ui->lineEditBirthday->text());
    query.bindValue(":personnel_id", ui->lineEditID->text());
    query.bindValue(":phone_num", ui->lineEditPhoneNum->text());
    query.bindValue(":race", ui->lineEditRace->text().section(' ', -1));
    query.bindValue(":degree", ui->lineEditDegree->text().section(' ', -1));
    query.bindValue(":health", ui->lineEditHealth->text().section(' ', -1));
    query.bindValue(":telephone_num", ui->lineEditTelephoneNum->text());
    query.bindValue(":edit_time", ui->lineEditTime->text());
    query.bindValue(":receipt", ui->lineEditReceipt->text().toUpper());
    query.bindValue(":workplace", ui->lineEditWorkPlace->text());
    query.bindValue(":province", ui->lineEditProvince->text());
    query.bindValue(":city", ui->lineEditCity->text());
    query.bindValue(":district", ui->lineEditDistrict->text());
    query.bindValue(":address", ui->lineEditAddress->text());
    query.bindValue(":postcode", ui->lineEditPostcode->text());
    query.bindValue(":graduate_time", ui->lineEditGraduateTime->text());
    query.bindValue(":graduate_school", ui->lineEditGraduateSchool->text());
    query.bindValue(":first_job_entry_time", ui->lineEditFirstJobDate->text());
    query.bindValue(":first_job_workplace", ui->lineEditFirstJobWorkspace->text());
    query.bindValue(":second_job_entry_time", ui->lineEditSecondJobDate->text());
    query.bindValue(":second_job_workplace", ui->lineEditSecondJobWorkspace->text());
    query.bindValue(":retirement_date", ui->lineEditRetirementDate->text());
    query.bindValue(":retirement_workplace", ui->lineEditRetirementWorkspace->text());
    query.bindValue(":year2start_learning_buddhism", ui->lineEditLearnBuddhismDate->text());
    query.bindValue(":years_of_learning_buddhism", ui->lineEditYearsOfLearningBuddhism->text());
    query.bindValue(":deep_understanding_of_dharma", ui->lineEditBuddhismLevel->text().section(' ', -1));
    query.bindValue(":reason2learning_dharma", ui->lineEditReasonToLearnDharma->text().section(' ', -1));
    query.bindValue(":nums_of_buddhism_book", ui->lineEditReadNumsOfBuddhismBook->text());
    query.bindValue(":easy2learn_buddhism_book", ui->lineEditEasyToLearnBuddhismBook->text());
    query.bindValue(":hard2read", ui->lineEditHardToLearnBuddhismBook->text());
    query.bindValue(":maxim", ui->lineEditMaximOfBuddhism->text());
    query.bindValue(":buddhist_disciples_of_family", ui->lineEditBuddhistDisciplesOfFamily->text());
    query.bindValue(":editor", ui->lineEditor->text());
    query.bindValue(":code", ui->lineEditCode->text());
    query.bindValue(":dbid", dbid);
    query.bindValue(":ipaddress", local_ip);
    dbid = 0;
    bool rt = query.exec();
    if (!rt) {
        // 输出错误到屏幕！[tbd] 这个得考虑到输入重复问题，下面的错误用户未必理解，所以得考虑到
        QString code = ui->lineEditCode->text();
        QString error_msg = query.lastError().text();
        if (error_msg.contains("Duplicate")) {
            QString message = query_duplicate(code);
            QMessageBox::information(this, "", message);
        } else {
            QMessageBox::information(this, "", QString("数据库错误，请联系开发运维人员：") + error_msg);
        }
        return false;
    }

    query.clear();

    return true;
}

QString Connect::query_duplicate(QString code)
{

    QSqlQuery query;
    query.prepare(
                " SELECT "
                "     `editor`, "
                "     `name` "
                " FROM "
                "     `zen_person` "
                " WHERE "
                "     `code` = :code"
                );

    query.bindValue(":code", code);
    query.exec();

    if (query.lastError().isValid()) {
        QMessageBox::critical(
                    this,
                    "数据库错误",
                    query.lastError().text()
                    );
        return "error";
    }

    query.next();

    return QString("确认一下您输入的皈依证号是否正确，如果正确，请找编辑人: [%1] 告诉他编辑的皈依号为: [%2] 皈依人名为: [%3] 的皈依号录入错误，请他改正后方可录入。")
            .arg(query.value(0).toString())
            .arg(code)
            .arg(query.value(1).toString());
}

bool Connect::clear_lineEdits() // [fixed xuefoxiaozu]
{
    if (!ui->actionJoinin->isEnabled()) {
        ui->lineEdit_Address->clear();
        ui->lineEdit_ApplyPlace->clear();
        ui->lineEdit_Birthday->clear();
        ui->lineEdit_City->clear();
        ui->lineEdit_ContractWay->clear();
        ui->lineEdit_District->clear();
        ui->lineEdit_Gender->clear();
        ui->lineEdit_Name->clear();
        ui->lineEdit_Order->clear(); // receipt field
        ui->lineEdit_Note->clear();
        ui->lineEdit_Phone->clear();
        ui->lineEdit_Province->clear();
        ui->lineEditLearnKind->clear();
        ui->lineEditLearnAddress->clear();
        return true;
    }

    ui->lineEditName->clear();
    ui->lineEditGender->clear();
    ui->lineEditJob->clear();
    ui->lineEditHobby->clear();
    ui->lineEditFName->clear();
    ui->lineEditBirthday->clear();
    ui->lineEditID->clear();
    ui->lineEditPhoneNum->clear();
    ui->lineEditDegree->clear();
    ui->lineEditHealth->clear();
    ui->lineEditTelephoneNum->clear();
    ui->lineEditRace->setText("汉族");
    //ui->lineEditProvince->clear();
    ui->lineEditProvince->setText("北京市");
    ui->lineEditCode->clear();
    //ui->lineEditCity->clear();
    ui->lineEditReceipt->clear();
    ui->lineEditCity->setText("北京市");
    ui->lineEditWorkPlace->clear();
    ui->lineEditDistrict->clear();
    ui->lineEditAddress->clear();
    ui->lineEditPostcode->clear();
    ui->lineEditGraduateTime->clear();
    ui->lineEditGraduateSchool->clear();
    ui->lineEditFirstJobDate->clear();
    ui->lineEditFirstJobWorkspace->clear();
    ui->lineEditSecondJobDate->clear();
    ui->lineEditSecondJobWorkspace->clear();
    ui->lineEditRetirementDate->clear();
    ui->lineEditRetirementWorkspace->clear();
    ui->lineEditLearnBuddhismDate->clear();
    ui->lineEditYearsOfLearningBuddhism->clear();
    ui->lineEditBuddhismLevel->clear();
    ui->lineEditReasonToLearnDharma->clear();
    ui->lineEditReadNumsOfBuddhismBook->clear();
    ui->lineEditEasyToLearnBuddhismBook->clear();
    ui->lineEditHardToLearnBuddhismBook->clear();
    ui->lineEditMaximOfBuddhism->clear();
    ui->lineEditBuddhistDisciplesOfFamily->clear();
    return true;
}


bool Connect::complete_fields(QString name, QString value)
{
    // 学佛小组录入
    if (!ui->actionJoinin->isEnabled()) {
        QSqlQuery query;
        QString sql = QString("SELECT receipt, name, gender, birthday, phone_num,"
                              "telephone_num, learn_dharma_kinds, province, city,"
                              "district, address, if_apply_learn_place, notes,"
                              "others, id, learn_dharma_address from zen_person where %1 = '%2'"
                              ).arg(name, value);
        query.exec(sql);
        qDebug() << sql;

        if (query.size() == 0) {
            //ui->lineEdit_Order->setFocus();
            return true;
        }

        if (query.size() > 1) {
            // tableView2 choose mode [tbd]
            query_by_name_or_phone(value);
            qDebug() << "Here we are!!!";
            return true;
        }

        while(query.next()) {
            ui->lineEdit_Order->setText(query.value(0).toString());
            ui->lineEdit_Name->setText(query.value(1).toString());
            ui->lineEdit_Gender->setText(query.value(2).toString());
            ui->lineEdit_Birthday->setText(query.value(3).toString());
            ui->lineEdit_Phone->setText(query.value(4).toString());
            ui->lineEdit_ContractWay->setText(query.value(5).toString());
            ui->lineEditLearnKind->setText(query.value(6).toString());
            ui->lineEdit_Province->setText(query.value(7).toString());
            ui->lineEdit_City->setText(query.value(8).toString());
            ui->lineEdit_District->setText(query.value(9).toString());
            ui->lineEdit_Address->setText(query.value(10).toString());
            ui->lineEdit_ApplyPlace->setText(query.value(11).toString());
            ui->lineEdit_Note->setText(query.value(12).toString());
            //ui->lineEditOthers->setText(query.value(13).toString());
            dbid = query.value(14).toInt();
            ui->lineEditLearnAddress->setText(query.value(15).toString());
        }
        return true;
    }

    QSqlQuery query;
    QString sql = QString(
               " SELECT "
               "    name, "
               "    gender, "
               "    job, "
               "    hobby, "
               "    fname, "
               "    birthday, "
               "    personnel_id, "
               "    phone_num, "
               "    race, "
               "    degree, "
               "    health, "
               "    telephone_num, "
               "    edit_time, "
               "    receipt, "
               "    workplace, "
               "    province, "
               "    city, "
               "    district, "
               "    address, "
               "    postcode, "
               "    graduate_time, "
               "    graduate_school, "
               "    first_job_entry_time, "
               "    first_job_workplace, "
               "    second_job_entry_time, "
               "    second_job_workplace, "
               "    retirement_date, "
               "    retirement_workplace, "
               "    year2start_learning_buddhism, "
               "    years_of_learning_buddhism, "
               "    deep_understanding_of_dharma, "
               "    reason2learning_dharma, "
               "    nums_of_buddhism_book, "
               "    easy2learn_buddhism_book, "
               "    hard2read, "
               "    maxim, "
               "    buddhist_disciples_of_family, "
               "    editor, "
               "    code, "
               "    id "
               " FROM "
               "    zen_person "
               " WHERE "
               "    %1 = '%2'"
            ).arg(name).arg(value);

    query.exec(sql);

    while(query.next()) {
        dbid = query.value(39).toInt();
        ui->lineEditName->setText(query.value(0).toString());
        ui->lineEditGender->setText(query.value(1).toString());
        ui->lineEditJob->setText(query.value(2).toString());
        ui->lineEditHobby->setText(query.value(3).toString());
        ui->lineEditFName->setText(query.value(4).toString());
        ui->lineEditBirthday->setText(query.value(5).toString());
        ui->lineEditID->setText(query.value(6).toString());
        ui->lineEditPhoneNum->setText(query.value(7).toString());
        ui->lineEditRace->setText(query.value(8).toString());
        ui->lineEditDegree->setText(query.value(9).toString());
        ui->lineEditHealth->setText(query.value(10).toString());
        ui->lineEditTelephoneNum->setText(query.value(11).toString());
        ui->lineEditTime->setText(query.value(12).toString());
        ui->lineEditReceipt->setText(query.value(13).toString().toUpper());
        ui->lineEditWorkPlace->setText(query.value(14).toString());
        ui->lineEditProvince->setText(query.value(15).toString());
        ui->lineEditCity->setText(query.value(16).toString());
        ui->lineEditDistrict->setText(query.value(17).toString());
        ui->lineEditAddress->setText(query.value(18).toString());
        ui->lineEditPostcode->setText(query.value(19).toString());
        ui->lineEditGraduateTime->setText(query.value(20).toString());
        ui->lineEditGraduateSchool->setText(query.value(21).toString());
        ui->lineEditFirstJobDate->setText(query.value(22).toString());
        ui->lineEditFirstJobWorkspace->setText(query.value(23).toString());
        ui->lineEditSecondJobDate->setText(query.value(24).toString());
        ui->lineEditSecondJobWorkspace->setText(query.value(25).toString());
        ui->lineEditRetirementDate->setText(query.value(26).toString());
        ui->lineEditRetirementWorkspace->setText(query.value(27).toString());
        ui->lineEditLearnBuddhismDate->setText(query.value(28).toString());
        ui->lineEditYearsOfLearningBuddhism->setText(query.value(29).toString());
        ui->lineEditBuddhismLevel->setText(query.value(30).toString());
        ui->lineEditReasonToLearnDharma->setText(query.value(31).toString());
        ui->lineEditReadNumsOfBuddhismBook->setText(query.value(32).toString());
        ui->lineEditEasyToLearnBuddhismBook->setText(query.value(33).toString());
        ui->lineEditHardToLearnBuddhismBook->setText(query.value(34).toString());
        ui->lineEditMaximOfBuddhism->setText(query.value(35).toString());
        ui->lineEditBuddhistDisciplesOfFamily->setText(query.value(36).toString());
        ui->lineEditor->setText(query.value(37).toString());
        ui->lineEditCode->setText(query.value(38).toString());
    }
    query.clear();
    return true;
}

void Connect::append_items2_tableView()
{
    if (!ui->actionJoinin->isEnabled()) {
        QString order = ui->lineEdit_Order->text(); // use receipt field instead
        QString name = ui->lineEdit_Name->text();
        QString num = ui->lineEdit_Phone->text();
        QString phone = QString("%1 %2 %3").arg(num.mid(0, 3)).arg(num.mid(3, 4)).arg(num.mid(7, 4));
        QString contractWay = ui->lineEdit_ContractWay->text();
        QString kinds = ui->lineEditLearnKind->text();
        QString gender = ui->lineEdit_Gender->text();
        QString birth = ui->lineEdit_Birthday->text();
        QString province = ui->lineEdit_Province->text();
        QString city = ui->lineEdit_City->text();
        QString district = ui->lineEdit_District->text();
        QString address = ui->lineEdit_Address->text();
        QString if_apply_place = ui->lineEdit_ApplyPlace->text();
        QString notes = ui->lineEdit_Note->text();

        QList <QStandardItem*> standardItemList;
        QStandardItem *orderItem = new QStandardItem(order);
        QStandardItem *nameItem = new QStandardItem(name);
        QStandardItem *phoneItem = new QStandardItem(phone);
        QStandardItem *contractWayItem = new QStandardItem(contractWay);
        QStandardItem *kindsItem = new QStandardItem(kinds);
        QStandardItem *genderItem = new QStandardItem(gender);
        QStandardItem *birthItem = new QStandardItem(birth);
        QStandardItem *provinceItem = new QStandardItem(province);
        QStandardItem *cityItem = new QStandardItem(city);
        QStandardItem *districtItem = new QStandardItem(district);
        QStandardItem *addressItem = new QStandardItem(address);
        QStandardItem *if_apply_placeItem = new QStandardItem(if_apply_place);
        QStandardItem *notesItem = new QStandardItem(notes);

        viewModel->appendRow(
                    standardItemList
                    << orderItem
                    << nameItem
                    << phoneItem
                    << contractWayItem
                    << kindsItem
                    << genderItem
                    << birthItem
                    << provinceItem
                    << cityItem
                    << districtItem
                    << addressItem
                    << if_apply_placeItem
                    << notesItem
                    );
    } else {
        QString name = ui->lineEditName->text();
        QString num = ui->lineEditPhoneNum->text();
        QString phone = QString("%1 %2 %3").arg(num.mid(0, 3)).arg(num.mid(3, 4)).arg(num.mid(7, 4));
        QString receipt = ui->lineEditReceipt->text().toUpper();
        QString code = ui->lineEditCode->text();
        QString personnel_id = ui->lineEditID->text();
        QString personId = QString("%1 %2 %3 %4 %5").arg(personnel_id.mid(0, 3)).arg(personnel_id.mid(3, 3)).arg(personnel_id.mid(6, 4)).arg(personnel_id.mid(10, 4)).arg(personnel_id.mid(14, 4));
        QString tel = ui->lineEditTelephoneNum->text();
        QString province = ui->lineEditProvince->text();
        QString city = ui->lineEditCity->text();
        QString district = ui->lineEditDistrict->text();
        QString address = ui->lineEditAddress->text();

        QList <QStandardItem*> standardItemList;
        QStandardItem *nameItem = new QStandardItem(name);
        QStandardItem *phoneItem = new QStandardItem(phone);
        QStandardItem *receiptItem = new QStandardItem(receipt);
        QStandardItem *codeItem = new QStandardItem(code);
        QStandardItem *personIdItem = new QStandardItem(personId);
        QStandardItem *telItem = new QStandardItem(tel);
        QStandardItem *provinceItem = new QStandardItem(province);
        QStandardItem *cityItem = new QStandardItem(city);
        QStandardItem *districtItem = new QStandardItem(district);
        QStandardItem *addressItem = new QStandardItem(address);


        viewModel->appendRow(
                    standardItemList
                    << nameItem
                    << phoneItem
                    << receiptItem
                    << codeItem
                    << personIdItem
                    << telItem
                    << provinceItem
                    << cityItem
                    << districtItem
                    << addressItem
                    );
    }
}

void Connect::closeEvent(QCloseEvent *event)
{
    QMessageBox::StandardButton ask = QMessageBox::question(this, "connect", "你确定要关闭程序?\n", QMessageBox
                                                            ::Cancel|QMessageBox::Yes);
    if (ask != QMessageBox::Yes) {
        event->ignore();
        qDebug() << "ignore close event";
    } else {
        //db.close();
        qDebug() << "db connection is closed";
    }
}

void Connect::get_local_ip()
{
    QString tmp_ip;
    QList<QHostAddress> list = QNetworkInterface::allAddresses();
    foreach (QHostAddress address, list)
    {
        if(address.protocol() == QAbstractSocket::IPv4Protocol)
            tmp_ip = address.toString();
        if (tmp_ip.startsWith(server_ip.section(".", 0, 2)))
            local_ip = tmp_ip;
    }
}

void Connect::excelen(QString fileName) //[new]
{
#define FIELDS_COUNT 12
    QXlsx::Document xlsx;
    QSqlQuery query;
    query.prepare(
                " SELECT"
                "   `name`,"
                "   `gender`,"
                "   `birthday`,"
                "   `phone_num`,"
                "   `telephone_num`,"
                "   `learn_dharma_kinds`,"
                "   `province`,"
                "   `city`,"
                "   `district`,"
                "   `address`,"
                "   `if_apply_learn_place`,"
                "   `notes`"
                " FROM "
                "   `zen_person`"
                " WHERE"
                "   `others` IS NOT NULL "
                " AND learn_dharma_kinds like '%外语%' "
                );
    query.exec();

    QString names[] = {
        "姓名",
        "性别",
        "出生年月",
        "手机号码",
        "其他联系方式",
        "学佛小组种类",
        "省",
        "市",
        "区(县)",
        "详细地址",
        "是否提供场地",
        "备注"
    };

    int lens[] = {
        10,
        5,
        11,
        15,
        16,
        20,
        10,
        12,
        12,
        35,
        12,
        50
    };

    char start = 'A';
    for(int i = 0; i < FIELDS_COUNT; i++) {
        xlsx.setColumnWidth(i+1, lens[i]);
    }

    for(int i = 0; i < FIELDS_COUNT; i++) {
        xlsx.write(QString("%1%2").arg(char(start+i)).arg("1"), names[i]);
    }

    int j = 2;
    while(query.next()) {
        for(int i = 0; i < FIELDS_COUNT; i++) {
            xlsx.write(QString("%1%2").arg(char(start + i)).arg(j), query.value(i).toString());
        }
        j++;
    }

    xlsx.saveAs(fileName);
}
void Connect::excel(QString fileName) //[new]
{
#define FIELDS_COUNT 12
    QXlsx::Document xlsx;
    QSqlQuery query;
    query.prepare(
                " SELECT"
                "   `name`,"
                "   `gender`,"
                "   `birthday`,"
                "   `phone_num`,"
                "   `telephone_num`,"
                "   `learn_dharma_kinds`,"
                "   `province`,"
                "   `city`,"
                "   `district`,"
                "   `address`,"
                "   `if_apply_learn_place`,"
                "   `notes`"
                " FROM "
                "   `zen_person`"
                " WHERE"
                "   `others` IS NOT NULL "
                );
    query.exec();

    QString names[] = {
        "姓名",
        "性别",
        "出生年月",
        "手机号码",
        "其他联系方式",
        "学佛小组种类",
        "省",
        "市",
        "区(县)",
        "详细地址",
        "是否提供场地",
        "备注"
    };

    int lens[] = {
        10,
        5,
        11,
        15,
        16,
        20,
        10,
        12,
        12,
        35,
        12,
        50
    };

    char start = 'A';
    for(int i = 0; i < FIELDS_COUNT; i++) {
        xlsx.setColumnWidth(i+1, lens[i]);
    }

    for(int i = 0; i < FIELDS_COUNT; i++) {
        xlsx.write(QString("%1%2").arg(char(start+i)).arg("1"), names[i]);
    }

    int j = 2;
    while(query.next()) {
        for(int i = 0; i < FIELDS_COUNT; i++) {
            xlsx.write(QString("%1%2").arg(char(start + i)).arg(j), query.value(i).toString());
        }
        j++;
    }

    xlsx.saveAs(fileName);
}

void Connect::save_excel(QString fileName)
{
    int i = 2;
    int j;
    QXlsx::Document xlsx;
    QSqlQuery query;
    query.exec("select \
               name,\
               gender,\
               job,\
               hobby,\
               fname,\
               code,\
               birthday,\
               personnel_id,\
               phone_num,\
               race,\
               degree,\
               health,\
               telephone_num,\
               edit_time,\
               receipt,\
               workplace,\
               province,\
               city,\
               district,\
               address,\
               postcode,\
               graduate_time,\
               graduate_school,\
               first_job_entry_time,\
               first_job_workplace,\
               second_job_entry_time,\
               second_job_workplace,\
               retirement_date,\
               retirement_workplace,\
               year2start_learning_buddhism,\
               years_of_learning_buddhism,\
               deep_understanding_of_dharma,\
               reason2learning_dharma,\
               nums_of_buddhism_book,\
               easy2learn_buddhism_book,\
               hard2read,\
               maxim,\
               buddhist_disciples_of_family,\
               editor,\
               others,\
               learn_dharma_kinds,\
               learn_dharma_address\
               from zen_person order by receipt");


    {
        xlsx.setColumnWidth(1, 10); // 姓名
        xlsx.setColumnWidth(2, 5); // 性别
        xlsx.setColumnWidth(3, 20); // 职业
        xlsx.setColumnWidth(4, 20); // 特长
        xlsx.setColumnWidth(5, 8); // 法名
        xlsx.setColumnWidth(6, 16); // 皈依证编号
        xlsx.setColumnWidth(7, 11); // 出生年月
        xlsx.setColumnWidth(8, 20); // 身份证号码
        xlsx.setColumnWidth(9, 15); // 手机号码
        xlsx.setColumnWidth(10, 15); // 民族
        xlsx.setColumnWidth(11, 10); // 文化程度
        xlsx.setColumnWidth(12, 10); // 健康状况
        xlsx.setColumnWidth(13, 14); // 固定电话
        xlsx.setColumnWidth(14, 11); // 填表时间
        xlsx.setColumnWidth(15, 10); // 收据编号
        xlsx.setColumnWidth(16, 25); // 工作单位
        xlsx.setColumnWidth(17, 10); // 省
        xlsx.setColumnWidth(18, 12); // 市
        xlsx.setColumnWidth(19, 12); // 区县
        xlsx.setColumnWidth(20, 35); // 通讯地址
        xlsx.setColumnWidth(21, 10); // 邮政编码
        xlsx.setColumnWidth(22, 11); // 毕业时间
        xlsx.setColumnWidth(23, 20); // 毕业学校
        xlsx.setColumnWidth(24, 11); // 工作入职时间一
        xlsx.setColumnWidth(25, 25); // 工作单位一
        xlsx.setColumnWidth(26, 10); // 工作入职时间二
        xlsx.setColumnWidth(27, 25); // 工作单位二
        xlsx.setColumnWidth(28, 11); // 退休时间
        xlsx.setColumnWidth(29, 25); // 退休单位
        xlsx.setColumnWidth(30, 10); // 学佛时间
        xlsx.setColumnWidth(31, 10); // 学佛时长
        xlsx.setColumnWidth(32, 20); // 是否深刻了解佛法
        xlsx.setColumnWidth(33, 25); // 以何因缘接触佛法
        xlsx.setColumnWidth(34, 13); // 所读经典部数
        xlsx.setColumnWidth(35, 20); // 认为易学经典
        xlsx.setColumnWidth(36, 20); // 认为不易学经典
        xlsx.setColumnWidth(37, 50); // 感触最深的一句话
        xlsx.setColumnWidth(38, 20); // 家庭三宝成员

        xlsx.write("A1", "姓名");
        xlsx.write("B1", "性别");
        xlsx.write("C1", "职业");
        xlsx.write("D1", "特长");
        xlsx.write("E1", "法名");
        xlsx.write("F1", "皈依证编号");
        xlsx.write("G1", "出生年月");
        xlsx.write("H1", "身份证号");
        xlsx.write("I1", "手机号码");
        xlsx.write("J1", "民族");
        xlsx.write("K1", "文化程度");
        xlsx.write("L1", "健康状况");
        xlsx.write("M1", "固定电话");
        xlsx.write("N1", "填表时间");
        xlsx.write("O1", "收据编号");
        xlsx.write("P1", "工作单位");
        xlsx.write("Q1", "省(直辖市/自治区)");
        xlsx.write("R1", "市");
        xlsx.write("S1", "区/县");
        xlsx.write("T1", "通讯地址");
        xlsx.write("U1", "邮政编码");
        xlsx.write("V1", "毕业时间");
        xlsx.write("W1", "毕业学校");
        xlsx.write("X1", "工作入职时间1");
        xlsx.write("Y1", "工作单位1");
        xlsx.write("Z1", "工作入职时间2");
        xlsx.write("AA1", "工作入职单位2");
        xlsx.write("AB1", "退休时间");
        xlsx.write("AC1", "退休单位");
        xlsx.write("AD1", "学佛时间");
        xlsx.write("AE1", "学佛时长");
        xlsx.write("AF1", "是否深刻理解佛法");
        xlsx.write("AG1", "以何因缘接触佛法");
        xlsx.write("AH1", "所读经典部数");
        xlsx.write("AI1", "认为易学经典");
        xlsx.write("AJ1", "认为不易学经典");
        xlsx.write("AK1", "感悟最深的一句话");
        xlsx.write("AL1", "家庭成员三宝弟子");

    }
    // excel A-Z, AA-AZ, BA-BZ...
    // 总共42列，所以下面代码是可以的，再多就不能这么写了，我图省事就这么写了！
    while(query.next()) {
        char start = 'A';
        char ne = 'A';
        for(j = 0; j < 38; j++) {
            if (j < 26) {
                xlsx.write(QString("%1%2").arg(start).arg(i), query.value(j).toString());
                start++;
            } else {
                xlsx.write(QString("A%1%2").arg(ne).arg(i), query.value(j).toString());
                ne++;
            }
        }
        i++;
    }
    query.clear();
    xlsx.saveAs(fileName);
}

void Connect::on_pushButtonExport_clicked()
{
    if (!test_if_connected()) return;
    QString fileName;
    fileName = QFileDialog::getSaveFileName(this, "打开保存文件路径", "", "xlsx (*.xlsx)");
    QString fileName_en = fileName.section('.', 0,0) + QString(".en.xlsx");
    if (fileName.isNull()) return;

    if (!ui->actionJoinin->isEnabled()) {
        excel(fileName);
        excelen(fileName_en);
    } else
        save_excel(fileName);
}


void Connect::on_pushButton_clicked()
{
    if (!test_if_connected()) return;
    init_and_append_items2_tableView();
    // ui->tableView_2->hide();
    qDebug() << "dbid" << dbid;
}

bool Connect::init_and_append_items2_tableView_check()
{
    QString receipt = ui->lineEditQuery->text().trimmed().toUpper();
    QSqlQuery query;

    if (!ui->actionJoinin->isEnabled()) {
        QString others = ui->lineEditOthers->text().trimmed();
        if (others.isEmpty()) {
            append_model_data(0, "");
        } else {
            QString where = QString(" WHERE `others` = (select others from zen_person where receipt = '%1') AND `mark` = 0").arg(receipt);
            append_model_data(0, where);
        }
       return true;
    }

    query.prepare(
                " SELECT "
                "    `name`, "
                "    `phone_num`, "
                "    `receipt`, "
                "    `code`, "
                "    `personnel_id`, "
                "    `telephone_num`, "
                "    `province`, "
                "    `city`, "
                "    `district`, "
                "    `address`, "
                "    `editor` "
                " FROM "
                "    `zen_person` "
                " WHERE editor = (select editor from zen_person where receipt = :receipt) and `mark` = 0 "
                " ORDER BY "
                "    `receipt` DESC "
                );

    query.bindValue(":receipt", receipt);

    query.exec();

    qDebug() << query.lastQuery();

    set_old_model_view(); // 清除所有行，重新设置标题

    QString editor = NULL;

    while(query.next()) {
        QString name = query.value(0).toString();
        qDebug() << "NAME: " << name;
        QString num = query.value(1).toString();
        QString phone = QString("%1 %2 %3").arg(num.mid(0, 3)).arg(num.mid(3, 4)).arg(num.mid(7, 4));
        QString receipt = query.value(2).toString();
        QString code = query.value(3).toString();
        QString person = query.value(4).toString();
        QString personId = QString("%1 %2 %3 %4 %5").arg(person.mid(0, 3)).arg(person.mid(3, 3)).arg(person.mid(6, 4)).arg(person.mid(10, 4)).arg(person.mid(14, 4));
        QString tel = query.value(5).toString();
        QString province = query.value(6).toString();
        QString city = query.value(7).toString();
        QString district = query.value(8).toString();
        QString address = query.value(9).toString();
        editor = query.value(10).toString();

        QList <QStandardItem*> standardItemList;
        QStandardItem *nameItem = new QStandardItem(name);
        QStandardItem *phoneItem = new QStandardItem(phone);
        QStandardItem *receiptItem = new QStandardItem(receipt);
        QStandardItem *codeItem = new QStandardItem(code);
        QStandardItem *personnelIDItem = new QStandardItem(personId);
        QStandardItem *telItem = new QStandardItem(tel);
        QStandardItem *provinceItem = new QStandardItem(province);
        QStandardItem *cityItem = new QStandardItem(city);
        QStandardItem *districtItem = new QStandardItem(district);
        QStandardItem *addressItem = new QStandardItem(address);


        viewModel->appendRow(
                    standardItemList
                    << nameItem
                    << phoneItem
                    << receiptItem
                    << codeItem
                    << personnelIDItem
                    << telItem
                    << provinceItem
                    << cityItem
                    << districtItem
                    << addressItem
                    );
    }

    QString message = QString("<strong><i>当前校对的编辑人姓名: <font color=red>%1</font></i></strong>").arg(editor);
    status_label->setText(message);
    query.clear();
    return true;
}

bool Connect::init_and_append_items2_tableView()
{
    QSqlQuery query;

    if (!ui->actionJoinin->isEnabled()) {
        QString others = ui->lineEditOthers->text().trimmed();
        if (others.isEmpty()) {
            append_model_data(0, "");
        } else {
            QString where = QString(" WHERE `others` = '%1' AND `mark` = 0").arg(others);
            append_model_data(0, where);
        }
       return true;
    }

    QString editor = ui->lineEditor->text().trimmed();
    qDebug() << "editor: " << editor;

    if (editor.isEmpty()) {
        query.prepare(
                    " SELECT "
                    "    `name`, "
                    "    `phone_num`, "
                    "    `receipt`, "
                    "    `code`, "
                    "    `personnel_id`, "
                    "    `telephone_num`, "
                    "    `province`, "
                    "    `city`, "
                    "    `district`, "
                    "    `address` "
                    " FROM "
                    "    `zen_person` "
                    " ORDER BY "
                    "    `receipt` DESC "
                    );
    } else {
         query.prepare(
                    " SELECT "
                    "   `name`, "
                    "   `phone_num`, "
                    "   `receipt`,"
                    "   `code`, "
                    "   `personnel_id`, "
                    "   `telephone_num`, "
                    "   `province`, "
                    "   `city`, "
                    "   `district`, "
                    "   `address` "
                    " FROM "
                    "   `zen_person` "
                    " WHERE "
                    "   `editor` = :editor "
                    "    AND `mark` = 0 "
                    " ORDER BY "
                    "    `receipt` "
                    );
         query.bindValue(":editor", editor);
    }

    query.exec();

    qDebug() << query.lastQuery();

    set_old_model_view(); // 清除所有行，重新设置标题

    while(query.next()) {
        QString name = query.value(0).toString();
        qDebug() << "NAME: " << name;
        QString num = query.value(1).toString();
        QString phone = QString("%1 %2 %3").arg(num.mid(0, 3)).arg(num.mid(3, 4)).arg(num.mid(7, 4));
        QString receipt = query.value(2).toString();
        QString code = query.value(3).toString();
        QString person = query.value(4).toString();
        QString personId = QString("%1 %2 %3 %4 %5").arg(person.mid(0, 3)).arg(person.mid(3, 3)).arg(person.mid(6, 4)).arg(person.mid(10, 4)).arg(person.mid(14, 4));
        QString tel = query.value(5).toString();
        QString province = query.value(6).toString();
        QString city = query.value(7).toString();
        QString district = query.value(8).toString();
        QString address = query.value(9).toString();

        QList <QStandardItem*> standardItemList;
        QStandardItem *nameItem = new QStandardItem(name);
        QStandardItem *phoneItem = new QStandardItem(phone);
        QStandardItem *receiptItem = new QStandardItem(receipt);
        QStandardItem *codeItem = new QStandardItem(code);
        QStandardItem *personnelIDItem = new QStandardItem(personId);
        QStandardItem *telItem = new QStandardItem(tel);
        QStandardItem *provinceItem = new QStandardItem(province);
        QStandardItem *cityItem = new QStandardItem(city);
        QStandardItem *districtItem = new QStandardItem(district);
        QStandardItem *addressItem = new QStandardItem(address);


        viewModel->appendRow(
                    standardItemList
                    << nameItem
                    << phoneItem
                    << receiptItem
                    << codeItem
                    << personnelIDItem
                    << telItem
                    << provinceItem
                    << cityItem
                    << districtItem
                    << addressItem
                    );
    }

    query.clear();
    return true;
}


bool Connect::init_db()
{
    // mysql
    db = QSqlDatabase::addDatabase("QMYSQL");
    db.setHostName(server_ip);
    db.setDatabaseName("connect");
    db.setUserName("connect");
    db.setPassword("connect");
    if(!db.open()) {
        QMessageBox::critical(this, "数据库错误", db.lastError().text());
        return false;
    }

    if_connected = true;
    create_table();
    return true;
}

bool Connect::create_table()
{
    QString  zen_person =
            "CREATE TABLE IF NOT EXISTS `zen_person` ("
            "  `id` int(11) NOT NULL AUTO_INCREMENT,"
            "  `name` varchar(32) DEFAULT NULL,"
            "  `gender` varchar(10) DEFAULT NULL,"
            "  `job` varchar(32) DEFAULT NULL,"
            "  `hobby` varchar(64) DEFAULT NULL,"
            "  `fname` varchar(32) DEFAULT NULL,"
            "  `birthday` varchar(32) DEFAULT NULL,"
            "  `personnel_id` varchar(32) DEFAULT NULL,"
            "  `phone_num` varchar(20) DEFAULT NULL,"
            "  `race` varchar(20) DEFAULT NULL,"
            "  `degree` varchar(32) DEFAULT NULL,"
            "  `health` varchar(32) DEFAULT NULL,"
            "  `telephone_num` varchar(20) DEFAULT NULL,"
            "  `edit_time` varchar(32) DEFAULT NULL,"
            "  `receipt` varchar(32) DEFAULT NULL,"
            "  `workplace` varchar(64) DEFAULT NULL,"
            "  `province` varchar(64) DEFAULT NULL,"
            "  `city` varchar(64) DEFAULT NULL,"
            "  `district` varchar(64) DEFAULT NULL,"
            "  `address` varchar(64) DEFAULT NULL,"
            "  `postcode` varchar(64) DEFAULT NULL,"
            "  `graduate_time` varchar(32) DEFAULT NULL,"
            "  `graduate_school` varchar(64) DEFAULT NULL,"
            "  `first_job_entry_time` varchar(64) DEFAULT NULL,"
            "  `first_job_workplace` varchar(64) DEFAULT NULL,"
            "  `second_job_entry_time` varchar(32) DEFAULT NULL,"
            "  `second_job_workplace` varchar(64) DEFAULT NULL,"
            "  `retirement_date` varchar(32) DEFAULT NULL,"
            "  `retirement_workplace` varchar(64) DEFAULT NULL,"
            "  `year2start_learning_buddhism` varchar(32) DEFAULT NULL,"
            "  `years_of_learning_buddhism` varchar(10) DEFAULT NULL,"
            "  `deep_understanding_of_dharma` varchar(64) DEFAULT NULL,"
            "  `reason2learning_dharma` varchar(128) DEFAULT NULL,"
            "  `nums_of_buddhism_book` varchar(10) DEFAULT NULL,"
            "  `easy2learn_buddhism_book` varchar(128) DEFAULT NULL,"
            "  `hard2read` varchar(128) DEFAULT NULL,"
            "  `maxim` varchar(128) DEFAULT NULL,"
            "  `buddhist_disciples_of_family` varchar(128) DEFAULT NULL,"
            "  `editor` varchar(128) DEFAULT NULL,"
            "  `others` varchar(128) DEFAULT NULL,"
            "  `learn_dharma_kinds` varchar(64) DEFAULT NULL,"
            "  `learn_dharma_address` varchar(128) DEFAULT NULL,"
            "  `code` varchar(64) DEFAULT NULL,"
            "  `if_apply_learn_place` varchar(10) DEFAULT '否',"
            "  `notes` varchar(128) DEFAULT NULL,"
            "  `mark` tinyint(1) NOT NULL DEFAULT '0',"
            "  `ipaddress` varchar(32) DEFAULT NULL,"
            "  `comes` tinyint NOT NULL DEFAULT '0' comment '来自手机或者笔记本录入, 0: 笔记本, 1: 手机',"
            "  PRIMARY KEY (`id`),"
            "  UNIQUE KEY `code` (`code`)"
            ") ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8";

    QSqlQuery query;

    QString admin =
            "CREATE TABLE IF NOT EXISTS `admin` ("
            "  `stat_name` varchar(31) DEFAULT NULL,"
            "  `stat` int(11) DEFAULT NULL"
            ") ENGINE=InnoDB DEFAULT CHARSET=utf8";

    query.exec(zen_person);
    query.exec(admin);

    if (query.lastError().isValid()) {
        qDebug() << query.lastError().text();
        QMessageBox::critical(this, "数据库错误", query.lastError().text());
        query.clear();
        return false;
    }

    query.clear();

    return true;
}

bool Connect::db_port_test()
{
    bool ret;
    QTcpSocket tsock;
    tsock.connectToHost(server_ip, 3306);
    ret = tsock.waitForConnected(1000);
    if (ret) tsock.close();
    return ret;
}

void Connect::on_pushButtonDatabase_pressed()
{
    hide_menu_and_button();

    if (!db.isOpen() && db_port_test()) {
        if (init_db()) {
            ui->pushButtonDatabase->setText("已连");
            ui->pushButtonDatabase->setDisabled(true);
        }
    } else {
        QMessageBox::critical(this, "无法连接数据库", "请在编辑菜单设置正确的数据库地址以及端口.");
        ui->pushButtonDatabase->setDown(false);
    }
}

void Connect::on_action_triggered()
{
    bool ok;
    QString text = QInputDialog::getText(this, "server address", "server address", QLineEdit::Normal, "192.168.1.5", &ok);
    if (ok && !text.isEmpty()) {
        server_ip = text;
    }
    get_local_ip();
    status_label->setText(QString(" Local Address: [%1], Server Address: [%2]").arg(local_ip, server_ip));
}

/*[tidy plan] */
bool Connect::append_model_data(int c, QString where)
{
    QSqlQuery query;

    if (!ui->actionJoinin->isEnabled()) {
        QString sql = QString(
                    " SELECT "
                    "     `receipt`, "
                    "     `name`, "
                    "     `phone_num`, "
                    "     `telephone_num`, "
                    "     `learn_dharma_kinds`, "
                    "     `gender`, "
                    "     `birthday`, "
                    "     `province`, "
                    "     `city`, "
                    "     `district`, "
                    "     `address`, "
                    "     `if_apply_learn_place`, "
                    "     `notes`, "
                    "     `id` "
                    " FROM "
                    "     `zen_person` "
                    ) + where;

        query.exec(sql);
        qDebug() << sql;

        if (query.lastError().isValid()) {
            QMessageBox::critical(
                        this,
                        "数据库错误",
                        query.lastError().text()
                        );
            return false;
        }

        set_new_model_view();
        while(query.next()) {
            QString order = query.value(0).toString();
            QString name = query.value(1).toString();
            QString phone = query.value(2).toString();
            QString tel = query.value(3).toString();
            QString kinds = query.value(4).toString();
            QString gender = query.value(5).toString();
            QString birth = query.value(6).toString();
            QString province = query.value(7).toString();
            QString city = query.value(8).toString();
            QString district = query.value(9).toString();
            QString address = query.value(10).toString();
            QString if_apply = query.value(11).toString();
            QString notes = query.value(12).toString();

            QList <QStandardItem*> standardItemList;
            QStandardItem *orderItem = new QStandardItem(order);
            QStandardItem *nameItem = new QStandardItem(name);
            QStandardItem *phoneItem = new QStandardItem(phone);
            QStandardItem *telItem= new QStandardItem(tel);
            QStandardItem *kindsItem = new QStandardItem(kinds);
            QStandardItem *genderItem = new QStandardItem(gender);
            QStandardItem *birthItem = new QStandardItem(birth);
            QStandardItem *provinceItem = new QStandardItem(province);
            QStandardItem *cityItem = new QStandardItem(city);
            QStandardItem *districtItem = new QStandardItem(district);
            QStandardItem *addressItem = new QStandardItem(address);
            QStandardItem *if_applyItem = new QStandardItem(if_apply);
            QStandardItem *notesItem = new QStandardItem(notes);

            if (c == 0) {
                viewModel->appendRow(standardItemList
                                     << orderItem
                                     << nameItem
                                     << phoneItem
                                     << telItem
                                     << kindsItem
                                     << genderItem
                                     << birthItem
                                     << provinceItem
                                     << cityItem
                                     << districtItem
                                     << addressItem
                                     << if_applyItem
                                     << notesItem
                                     );
            } else {
                viewModel_search->appendRow(standardItemList
                                            << orderItem
                                            << nameItem
                                            << phoneItem
                                            << telItem
                                            << kindsItem
                                            << genderItem
                                            << birthItem
                                            << provinceItem
                                            << cityItem
                                            << districtItem
                                            << addressItem
                                            << if_applyItem
                                            << notesItem
                                            );

            }
        }

        query.clear();

        return true;
    }

    QString sql = QString(
                " SELECT "
                "   `name`, "
                "   `phone_num`, "
                "   `receipt`, "
                "   `code`, "
                "   `province`, "
                "   `city`, "
                "   `district`, "
                "   `address` "
                " FROM "
                "   `zen_person` "
                " WHERE "
                ) + where;

    qDebug() << sql;
    query.exec(sql);

    if (query.lastError().isValid()) {
        QMessageBox::critical(this, "数据库错误", query.lastError().text());
        return false;
    }

    set_old_model_view();

    while(query.next()) {
        QString name = query.value(0).toString();
        QString phone = query.value(1).toString();
        QString receipt = query.value(2).toString();
        QString code = query.value(3).toString();
        QString province = query.value(4).toString();
        QString city = query.value(5).toString();
        QString district = query.value(6).toString();
        QString address = query.value(7).toString();

        QList <QStandardItem*> standardItemList;
        QStandardItem *nameItem = new QStandardItem(name);
        QStandardItem *phoneItem = new QStandardItem(phone);
        QStandardItem *receiptItem = new QStandardItem(receipt);
        QStandardItem *codeItem = new QStandardItem(code);
        QStandardItem *provinceItem = new QStandardItem(province);
        QStandardItem *cityItem = new QStandardItem(city);
        QStandardItem *districtItem = new QStandardItem(district);
        QStandardItem *addressItem = new QStandardItem(address);

        if (c == 0) {
        viewModel->appendRow(standardItemList
                                    << nameItem
                                    << phoneItem
                                    << receiptItem
                                    << codeItem
                                    << provinceItem
                                    << cityItem
                                    << districtItem
                                    << addressItem
                                    );
        } else {
         viewModel_search->appendRow(standardItemList
                                    << nameItem
                                    << phoneItem
                                    << receiptItem
                                    << codeItem
                                    << provinceItem
                                    << cityItem
                                    << districtItem
                                    << addressItem
                                    );

        }
    }

    query.clear();
    return true;
}

void Connect::query_by_name_or_phone(QString search_text)
{
    QSqlQuery query;
    query.prepare(
                " SELECT "
                "     `receipt`, "
                "     `name`, "
                "     `phone_num`, "
                "     `telephone_num`, "
                "     `learn_dharma_kinds`, "
                "     `gender`, "
                "     `birthday`, "
                "     `province`, "
                "     `city`, "
                "     `district`, "
                "     `address`, "
                "     `if_apply_learn_place`, "
                "     `notes`, "
                "     `id` "
                " FROM "
                "     `zen_person` "
                " WHERE "
                "     `name` = :abc "
                " OR "
                "   `phone_num` = :abc "
                );

    query.bindValue(":abc", search_text);
    query.exec();

    if (query.lastError().isValid()) {
        QMessageBox::critical(
                    this,
                    "数据库错误",
                    query.lastError().text()
                    );
        return;
    }

    set_new_second_model();
    while(query.next()) {
        QString order = query.value(0).toString();
        QString name = query.value(1).toString();
        QString phone = query.value(2).toString();
        QString tel = query.value(3).toString();
        QString kinds = query.value(4).toString();
        QString gender = query.value(5).toString();
        QString birth = query.value(6).toString();
        QString province = query.value(7).toString();
        QString city = query.value(8).toString();
        QString district = query.value(9).toString();
        QString address = query.value(10).toString();
        QString if_apply = query.value(11).toString();
        QString notes = query.value(12).toString();

        QList <QStandardItem*> standardItemList;
        QStandardItem *orderItem = new QStandardItem(order);
        QStandardItem *nameItem = new QStandardItem(name);
        QStandardItem *phoneItem = new QStandardItem(phone);
        QStandardItem *telItem= new QStandardItem(tel);
        QStandardItem *kindsItem = new QStandardItem(kinds);
        QStandardItem *genderItem = new QStandardItem(gender);
        QStandardItem *birthItem = new QStandardItem(birth);
        QStandardItem *provinceItem = new QStandardItem(province);
        QStandardItem *cityItem = new QStandardItem(city);
        QStandardItem *districtItem = new QStandardItem(district);
        QStandardItem *addressItem = new QStandardItem(address);
        QStandardItem *if_applyItem = new QStandardItem(if_apply);
        QStandardItem *notesItem = new QStandardItem(notes);

        viewModel_search->appendRow(standardItemList
                                    << orderItem
                                    << nameItem
                                    << phoneItem
                                    << telItem
                                    << kindsItem
                                    << genderItem
                                    << birthItem
                                    << provinceItem
                                    << cityItem
                                    << districtItem
                                    << addressItem
                                    << if_applyItem
                                    << notesItem
                                    );
    }

    show_search_table();
    query.clear();

    return;
}

void Connect::on_actionQueryAnyThing_triggered()
{
    if (!test_if_connected()) return;
    bool ok;
    QString text = QString("请输入姓名、手机号、皈依号进行查询，输入后按确定，或者点击 OK");
    QString search_text = QInputDialog::getText(this, "查询窗口", text, QLineEdit::Normal, "", &ok );

    if (!(ok && !search_text.isEmpty())) return;

    QSqlQuery query;
    search_text = search_text.trimmed();


    if (!ui->actionJoinin->isEnabled()) {
        query_by_name_or_phone(search_text);
        return;
    }

    query.prepare(
                " SELECT "
                "   `name`, "
                "   `phone_num`, "
                "   `receipt`, "
                "   `code`, "
                "   `personnel_id`, "
                "   `telephone_num`, "
                "   `province`, "
                "   `city`, "
                "   `district`, "
                "   `address` "
                " FROM "
                "   `zen_person` "
                " WHERE "
                "   `name` = :abc "
                " OR "
                "   `phone_num` = :abc "
                " OR "
                "   `receipt` = :abc "
                " OR "
                "   `code` = :abc "
                );
    query.bindValue(":abc", search_text);
    query.exec();

    if (query.lastError().isValid()) {
        QMessageBox::critical(this, "数据库错误", query.lastError().text());
        return;
    }

    set_old_second_model();
    while(query.next()) {
        QString name = query.value(0).toString();
        QString phone = query.value(1).toString();
        QString receipt = query.value(2).toString();
        QString code = query.value(3).toString();
        QString personnel_id = query.value(4).toString();
        QString telephone_num = query.value(5).toString();
        QString province = query.value(6).toString();
        QString city = query.value(7).toString();
        QString district = query.value(8).toString();
        QString address = query.value(9).toString();

        QList <QStandardItem*> standardItemList;
        QStandardItem *nameItem = new QStandardItem(name);
        QStandardItem *phoneItem = new QStandardItem(phone);
        QStandardItem *receiptItem = new QStandardItem(receipt);
        QStandardItem *codeItem = new QStandardItem(code);
        QStandardItem *personnel_idItem = new QStandardItem(personnel_id);
        QStandardItem *telephone_numItem = new QStandardItem(telephone_num);
        QStandardItem *provinceItem = new QStandardItem(province);
        QStandardItem *cityItem = new QStandardItem(city);
        QStandardItem *districtItem = new QStandardItem(district);
        QStandardItem *addressItem = new QStandardItem(address);

        viewModel_search->appendRow(standardItemList
                                    << nameItem
                                    << phoneItem
                                    << receiptItem
                                    << codeItem
                                    << personnel_idItem
                                    << telephone_numItem
                                    << provinceItem
                                    << cityItem
                                    << districtItem
                                    << addressItem
                                    );
    }

    show_search_table();

    query.clear();
}

void Connect::hide_search_table()
{
    ui->tableView_2->hide();
}

void Connect::show_search_table()
{
    ui->tableView_2->show();
}

void Connect::on_tableView_2_doubleClicked(const QModelIndex &index)
{
    int num;
    if(!ui->actionJoinin->isEnabled()) {
        num = 0;
    } else {
        num = 2;
    }

    QString receipt = index.sibling(index.row(), num).data().toString();
    if (modify_or_not()) {
        complete_fields("receipt", receipt);
        hide_search_table();
    } else return;
}

bool Connect::test_if_connected()
{
    if(!if_connected) {
        QMessageBox::critical(this, "没有连接数据库", "请先连接数据库");
        return false;
    }
    return true;
}

void Connect::on_actionDbBack_triggered()
{
    if (!test_if_connected()) return;
    bool retcode;
    QDateTime dt;
    QString current_dt = dt.currentDateTime().toString("yyyy_MM_dd_hh_mm_ss");
    qDebug() << current_dt;

    QString new_table_name = QString("zen_person_%1").arg(current_dt);
    QString sql = QString("create table %1 as select * from zen_person;").arg(new_table_name);
    QSqlQuery query;
    retcode = query.exec(sql);
    qDebug() << retcode;
    if (!retcode)
        QMessageBox::critical(this, "备份", QString("备份失败: %1").arg(query.lastError().text()));
    else
        QMessageBox::information(this, "备份", QString("备份数据到新表格成功: %1").arg(new_table_name));
}

void Connect::hide_menu_and_button()
{
    get_local_ip();
    if (local_ip.section('.', -1).toInt() > IP_TAIL_MAX) {
        ui->pushButtonExport->hide();
        //ui->pushButton->hide();
        ui->actionDbBack->setDisabled(true);
    }
}

void Connect::on_actionJoinin_triggered()
{
    int stat;
    if (local_ip.section('.', -1).toInt() > IP_TAIL_MAX) {
        QSqlQuery query;
        query.prepare("select stat from admin where stat_name = 'xuefoxiaozu'");
        query.exec();
        query.next();
        stat = query.value(0).toInt();
        if (stat == 0) {
            QMessageBox::information(this, "学佛小组录入", "管理员未允许学佛小组录入，请联系管理员，管理员点击后方可录入。");
            return;
        }
    }

    // 更新状态表，允许所有使用者录入学佛小组信息
    update_table("update `admin` set `stat` = 1 where `stat_name` = 'xuefoxiaozu'");
    update_table("update zen_person set mark = 0");
    ui->fgroupWidget->show();
    ui->joinWidget->hide();
    ui->actionJoinin->setDisabled(true);
    set_new_model_view();
    ui->lineEdit_ApplyPlace->setText("否");
}

void Connect::on_lineEdit_Name_returnPressed()
{
    qDebug() << "on line edit return pressed";
    complete_fields("name", ui->lineEdit_Name->text());
}

void Connect::on_lineEdit_Name_editingFinished()
{
    complete_fields("name", ui->lineEdit_Name->text());
}

void Connect::set_old_first_model()
{
    viewModel->clear();
    viewModel = new QStandardItemModel();
    viewModel->setHorizontalHeaderItem(0, new QStandardItem("姓名"));
    viewModel->setHorizontalHeaderItem(1, new QStandardItem("手机"));
    viewModel->setHorizontalHeaderItem(2, new QStandardItem("收据编号"));
    viewModel->setHorizontalHeaderItem(3, new QStandardItem("皈依证号"));
    viewModel->setHorizontalHeaderItem(4, new QStandardItem("身份证号码"));
    viewModel->setHorizontalHeaderItem(5, new QStandardItem("固定电话"));
    viewModel->setHorizontalHeaderItem(6, new QStandardItem("省"));
    viewModel->setHorizontalHeaderItem(7, new QStandardItem("市"));
    viewModel->setHorizontalHeaderItem(8, new QStandardItem("区/县"));
    viewModel->setHorizontalHeaderItem(9, new QStandardItem("村/街道"));

    ui->tableView->setModel(viewModel);
    ui->tableView->setEditTriggers(QAbstractItemView::NoEditTriggers);
}

void Connect::set_old_second_model()
{
    viewModel_search->clear();
    viewModel_search = new QStandardItemModel();
    viewModel_search->setHorizontalHeaderItem(0, new QStandardItem("姓名"));
    viewModel_search->setHorizontalHeaderItem(1, new QStandardItem("手机"));
    viewModel_search ->setHorizontalHeaderItem(2, new QStandardItem("收据编号"));
    viewModel_search->setHorizontalHeaderItem(3, new QStandardItem("皈依证号"));

    viewModel_search->setHorizontalHeaderItem(4, new QStandardItem("身份证号码"));
    viewModel_search->setHorizontalHeaderItem(5, new QStandardItem("固定电话"));
    viewModel_search->setHorizontalHeaderItem(6, new QStandardItem("省"));
    viewModel_search->setHorizontalHeaderItem(7, new QStandardItem("市"));
    viewModel_search->setHorizontalHeaderItem(8, new QStandardItem("区/县"));
    viewModel_search->setHorizontalHeaderItem(9, new QStandardItem("村/街道"));
    ui->tableView_2->setModel(viewModel_search);
    ui->tableView_2->setEditTriggers(QAbstractItemView::NoEditTriggers);
}

void Connect::set_old_model_view()
{
    set_old_first_model();
    set_old_second_model();
}

void Connect::set_new_first_model()
{
    viewModel->clear();
    viewModel = new QStandardItemModel();
    viewModel->setHorizontalHeaderItem(0, new QStandardItem("序号"));
    viewModel->setHorizontalHeaderItem(1, new QStandardItem("姓名"));
    viewModel->setHorizontalHeaderItem(2, new QStandardItem("手机"));
    viewModel->setHorizontalHeaderItem(3, new QStandardItem("其他联系方式"));
    viewModel->setHorizontalHeaderItem(4, new QStandardItem("小组种类"));
    viewModel->setHorizontalHeaderItem(5, new QStandardItem("性别"));
    viewModel->setHorizontalHeaderItem(6, new QStandardItem("生日"));
    viewModel->setHorizontalHeaderItem(7, new QStandardItem("省"));
    viewModel->setHorizontalHeaderItem(8, new QStandardItem("市"));
    viewModel->setHorizontalHeaderItem(9, new QStandardItem("区(县)"));
    viewModel->setHorizontalHeaderItem(10, new QStandardItem("街道(村)"));
    viewModel->setHorizontalHeaderItem(11, new QStandardItem("是否提供场地"));
    viewModel->setHorizontalHeaderItem(12, new QStandardItem("备注"));

    ui->tableView->setModel(viewModel);
    ui->tableView->setEditTriggers(QAbstractItemView::NoEditTriggers);
}

void Connect::set_new_second_model()
{
    viewModel_search->clear();
    viewModel_search = new QStandardItemModel();
    viewModel_search->setHorizontalHeaderItem(0, new QStandardItem("序号"));
    viewModel_search->setHorizontalHeaderItem(1, new QStandardItem("姓名"));
    viewModel_search->setHorizontalHeaderItem(2, new QStandardItem("手机"));
    viewModel_search->setHorizontalHeaderItem(3, new QStandardItem("其他联系方式"));
    viewModel_search->setHorizontalHeaderItem(4, new QStandardItem("小组种类"));
    viewModel_search->setHorizontalHeaderItem(5, new QStandardItem("性别"));
    viewModel_search->setHorizontalHeaderItem(6, new QStandardItem("生日"));
    viewModel_search->setHorizontalHeaderItem(7, new QStandardItem("省"));
    viewModel_search->setHorizontalHeaderItem(8, new QStandardItem("市"));
    viewModel_search->setHorizontalHeaderItem(9, new QStandardItem("区(县)"));
    viewModel_search->setHorizontalHeaderItem(10, new QStandardItem("街道(村)"));
    viewModel_search->setHorizontalHeaderItem(11, new QStandardItem("是否提供场地"));
    viewModel_search->setHorizontalHeaderItem(12, new QStandardItem("备注"));

    ui->tableView_2->setModel(viewModel_search);
    ui->tableView_2->setEditTriggers(QAbstractItemView::NoEditTriggers);
}

void Connect::set_new_model_view()
{
    set_new_first_model();
    set_new_second_model();
}

/* 根据皈依证号，自动判断性别 */
void Connect::on_lineEditCode_editingFinished()
{
    if (ui->lineEditCode->text().isEmpty()) {
        return;
    }

    QChar s = ui->lineEditCode->text().at(0).toLower();
    if (s == 'a') {
        ui->lineEditCode->setText(ui->lineEditCode->text().toUpper());
        ui->lineEditGender->setText("男");
        return;
    } else if (s == 'b') {
        ui->lineEditCode->setText(ui->lineEditCode->text().toUpper());
        ui->lineEditGender->setText("女");
        return;
    }

    ui->lineEditCode->setText("号码有误，请重新输入");
    ui->lineEditCode->selectAll();
    ui->lineEditCode->setFocus();
}

bool Connect::admin_init_all()
{
    // init db, create table, truncate table, save table
    update_table("TRUNCATE `zen_person`");
    update_table("drop table if exists `admin`");
    update_table("create table if not exists `admin` (`stat_name` varchar(32), `stat` int)");
    update_table("insert into `admin` (`stat_name`, `stat`) values ('xuefoxiaozu', 0)");
    return true;
}

bool Connect::update_table(QString upsql)
{
    QSqlQuery query;
    query.exec(upsql);
    if (query.lastError().isValid()) {
        query.clear();
        qDebug() << query.lastError().text();
        return false;
    }
    query.clear();
    return true;
}

/* Table view mouse right click menu functions */
void Connect::on_tableView_customContextMenuRequested(const QPoint &pos)
{
    // IP 不是管理员段的没有权限执行下面
    if (local_ip.section(".", -1).toInt() > IP_TAIL_MAX) {
        return;
    }

    int col;
    if (!ui->actionJoinin->isEnabled())
        col = 0;
    else
        col = 2;

    int row = ui->tableView->verticalHeader()->logicalIndexAt(pos);
    if (row < 0) return;
    grow = row;

    QMenu *popMenu = new QMenu(this);
    QString id = viewModel->index(row, col).data().toString();
    gsql = QString(
                " UPDATE "
                "   `zen_person` "
                " SET"
                "   `mark` = 1 "
                " WHERE"
                "   `receipt` = '%1' "
                ).arg(id);

    popMenu->addAction(ui->actionCheck);
    popMenu->exec(QCursor::pos());
}

void Connect::on_actionCheck_triggered()
{
    update_table(gsql);
    viewModel->removeRow(grow);
    grow = -1; // init to no used value
}

void Connect::on_actionInitDb_triggered()
{
    if (!test_if_connected()) return;
    QMessageBox::StandardButton ask
            = QMessageBox::question(this,
                                    "connect",
                                    "<font color=\"red\">危险!!!</font> 你确定初始化删除所有数据么? 如果确定请按 Yes，否则按 Cancel\n",
                                    QMessageBox::Cancel|QMessageBox::Yes);
    if (ask == QMessageBox::Yes) {
        admin_init_all();
        return;
    }
    return;
}

void Connect::on_actionAbort_triggered()
{
   // 在编辑模式下点击 放弃编辑 清空数据状态标记，还原数据到显示表格
    dbid = 0;
    QString edit_name;
    QString where;
    if (!ui->actionJoinin->isEnabled()) {
        edit_name = ui->lineEditOthers->text();
        qDebug() << "Here we are!";
        where = QString(" WHERE `others` = '%1'").arg(edit_name);
    } else {
        qDebug() << "We are here";
        edit_name = ui->lineEditor->text();
        where = QString(" `editor` = '%1' and mark = 0").arg(edit_name);
        edit_name = ui->lineEditor->text();
    }

    clear_lineEdits();
    if (!edit_name.isEmpty()) {
        append_model_data(0, where);
    }
}

void Connect::on_lineEdit_Phone_editingFinished()
{
    complete_fields("phone_num", ui->lineEdit_Phone->text());
}

void Connect::on_actionRevision_triggered()
{
    if (ui->lineEditQuery->isHidden()) {
        ui->lineEditQuery->show();
        ui->lineEditQuery->setFocus();
    } else {
        ui->lineEditQuery->hide();
    }
}

void Connect::on_lineEditQuery_returnPressed()
{
    init_and_append_items2_tableView_check();
    ui->lineEditQuery->hide();
}

void Connect::on_lineEditName_editingFinished()
{
    int len = ui->lineEditName->text().size();
    QString fname = QString("贤%1").arg(ui->lineEditName->text()[len-1]);
    ui->lineEditFName->setText(fname);
}



void Connect::on_lineEdit_Order_returnPressed()
{
    qDebug() << "Here we are";
    QString receipt = ui->lineEdit_Order->text();

    complete_fields("receipt", receipt);
}

