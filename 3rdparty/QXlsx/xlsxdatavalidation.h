#pragma once

#include <QList>
#include <QSharedDataPointer>
#include <QString>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QtGlobal>

#include "xlsxglobal.h"

class QXmlStreamReader;
class QXmlStreamWriter;

QT_BEGIN_NAMESPACE_XLSX

class Worksheet;
class CellRange;
class CellReference;

class DataValidationPrivate;
class DataValidation
{
public:
	enum class Type
	{
		None,
		Whole,
		Decimal,
		List,
		Date,
		Time,
		TextLength,
		Custom
	};

	enum class Operator
	{
		Between,
		NotBetween,
		Equal,
		NotEqual,
		LessThan,
		LessThanOrEqual,
		GreaterThan,
		GreaterThanOrEqual
	};

	enum class ErrorStyle
	{
		Stop,
		Warning,
		Information
	};

	DataValidation();
	DataValidation(Type type, Operator op = Operator::Between, const QString &formula1 = QString(), const QString &formula2 = QString(), bool allowBlank = false);
	DataValidation(const DataValidation &other);
	~DataValidation();

	Type validationType() const;
	Operator validationOperator() const;
	ErrorStyle errorStyle() const;
	QString formula1() const;
	QString formula2() const;
	bool allowBlank() const;
	QString errorMessage() const;
	QString errorMessageTitle() const;
	QString promptMessage() const;
	QString promptMessageTitle() const;
	bool isPromptMessageVisible() const;
	bool isErrorMessageVisible() const;
	QList<CellRange> ranges() const;

	void setType(Type type);
	void setOperator(Operator op);
	void setErrorStyle(ErrorStyle es);
	void setFormula1(const QString &formula);
	void setFormula2(const QString &formula);
	void setErrorMessage(const QString &error, const QString &title = QString());
	void setPromptMessage(const QString &prompt, const QString &title = QString());
	void setAllowBlank(bool enable);
	void setPromptMessageVisible(bool visible);
	void setErrorMessageVisible(bool visible);

	void addCell(const CellReference &cell);
	void addCell(int row, int col);
	void addRange(int firstRow, int firstCol, int lastRow, int lastCol);
	void addRange(const CellRange &range);

	DataValidation &operator=(const DataValidation &other);

	bool saveToXml(QXmlStreamWriter &writer) const;
	static DataValidation loadFromXml(QXmlStreamReader &reader);

private:
	QSharedDataPointer<DataValidationPrivate> d;
};

QT_END_NAMESPACE_XLSX
