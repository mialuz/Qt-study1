#include <QtGui>

#include "cell.h"

Cell::Cell()
{
    setDirty();
}

//在创建新的单元格时，调用clone用于克隆新的实例
QTableWidgetItem *Cell::clone() const
{
    return new Cell(*this);
}

void Cell::setData(int role, const QVariant &value)
{
    QTableWidgetItem::setData(role, value);
    if (role == Qt::EditRole)
        setDirty();
}

QVariant Cell::data(int role) const
{
    if (role == Qt::DisplayRole) {
        if (value().isValid()) {
            return value().toString();
        } else {
            return "####";
        }
    } else if (role == Qt::TextAlignmentRole) {
        if (value().type() == QVariant::String) {
            return int(Qt::AlignLeft | Qt::AlignVCenter);
        } else {
            return int(Qt::AlignRight | Qt::AlignVCenter);
        }
    } else {
        return QTableWidgetItem::data(role);
    }
}

//用来设置单元格中的公式
void Cell::setFormula(const QString &formula)
{
    setData(Qt::EditRole, formula);
}

//在spreadsheet::formula()中得到调用，重新获得该项的EditRole数据
QString Cell::formula() const
{
    return data(Qt::EditRole).toString();
}

//用来对该单元格的值强制进行重新计算
void Cell::setDirty()
{
    cacheIsDirty = true;
}

const QVariant Invalid;  //使用默认构造函数构造的QVariant是一个“无效”的变量

//value私有函数返回这个单元格的值
QVariant Cell::value() const
{
    //如果cacheIsDirty是true，就需要重新计算这个值
    if (cacheIsDirty) {
        cacheIsDirty = false;

        QString formulaStr = formula();
        //如果公式是由单引号开始的，那么这个单引号占用位置0,而值就是从位置1直到最后位置的一个字符串
        if (formulaStr.startsWith('\'')) {
            cachedValue = formulaStr.mid(1);
        } else if (formulaStr.startsWith('=')) {
        //如果公式是由等号开始的，值就是从位置1直到最后位置的一个字符串，并将它可能包含的任意空格全部移除
            cachedValue = Invalid;
            QString expr = formulaStr.mid(1);  
            expr.replace(" ", "");
            expr.append(QChar::Null);

            int pos = 0;
            cachedValue = evalExpression(expr, pos);   //计算表达式的值
            if (expr[pos] != QChar::Null)
                cachedValue = Invalid;
        } else {
            //不是单引号或等号开头的
            bool ok;
            double d = formulaStr.toDouble(&ok); //用于区分数字还是字符串
            if (ok) {
                cachedValue = d;  //数字
            } else {
                cachedValue = formulaStr;  //字符串
            }
        }
    }
    return cachedValue;
}

//返回一个电子制表软件表达式的值
QVariant Cell::evalExpression(const QString &str, int &pos) const
{
    QVariant result = evalTerm(str, pos);
    while (str[pos] != QChar::Null) {
        QChar op = str[pos];
        if (op != '+' && op != '-')
            return result;
        ++pos;

        QVariant term = evalTerm(str, pos);
        if (result.type() == QVariant::Double
                && term.type() == QVariant::Double) {
            if (op == '+') {
                result = result.toDouble() + term.toDouble();
            } else {
                result = result.toDouble() - term.toDouble();
            }
        } else {
            result = Invalid;
        }
    }
    return result;
}

QVariant Cell::evalTerm(const QString &str, int &pos) const
{
    QVariant result = evalFactor(str, pos);
    while (str[pos] != QChar::Null) {
        QChar op = str[pos];
        if (op != '*' && op != '/')
            return result;
        ++pos;

        QVariant factor = evalFactor(str, pos);
        if (result.type() == QVariant::Double
                && factor.type() == QVariant::Double) {
            if (op == '*') {
                result = result.toDouble() * factor.toDouble();
            } else {
                if (factor.toDouble() == 0.0) {
                    result = Invalid;
                } else {
                    result = result.toDouble() / factor.toDouble();
                }
            }
        } else {
            result = Invalid;
        }
    }
    return result;
}

QVariant Cell::evalFactor(const QString &str, int &pos) const
{
    QVariant result;
    bool negative = false;

    if (str[pos] == '-') {
        negative = true;
        ++pos;
    }

    if (str[pos] == '(') {
        ++pos;
        result = evalExpression(str, pos);
        if (str[pos] != ')')
            result = Invalid;
        ++pos;
    } else {
        QRegExp regExp("[A-Za-z][1-9][0-9]{0,2}");
        QString token;

        while (str[pos].isLetterOrNumber() || str[pos] == '.') {
            token += str[pos];
            ++pos;
        }

        if (regExp.exactMatch(token)) {
            int column = token[0].toUpper().unicode() - 'A';
            int row = token.mid(1).toInt() - 1;

            Cell *c = static_cast<Cell *>(
                              tableWidget()->item(row, column));
            if (c) {
                result = c->value();
            } else {
                result = 0.0;
            }
        } else {
            bool ok;
            result = token.toDouble(&ok);
            if (!ok)
                result = Invalid;
        }
    }

    if (negative) {
        if (result.type() == QVariant::Double) {
            result = -result.toDouble();
        } else {
            result = Invalid;
        }
    }
    return result;
}
