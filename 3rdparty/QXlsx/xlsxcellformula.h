#pragma once

#include "xlsxglobal.h"

#include <QExplicitlySharedDataPointer>

class QXmlStreamWriter;
class QXmlStreamReader;

QT_BEGIN_NAMESPACE_XLSX

class CellFormulaPrivate;
class CellRange;
class Worksheet;
class WorksheetPrivate;

class CellFormula
{
public:
	enum class Type
	{
		Normal,
		Array,
		DataTable,
		Shared
	};

	CellFormula();
	CellFormula(const char *formula, Type type = Type::Normal);
	CellFormula(const QString &formula, Type type = Type::Normal);
	CellFormula(const QString &formula, const CellRange &ref, Type type);
	CellFormula(const CellFormula &other);
	~CellFormula();

	CellFormula &operator=(const CellFormula &other);
	bool isValid() const;

	Type formulaType() const;
	QString formulaText() const;
	CellRange reference() const;
	int sharedIndex() const;

	bool operator==(const CellFormula &formula) const;
	bool operator!=(const CellFormula &formula) const;

	bool saveToXml(QXmlStreamWriter &writer) const;
	bool loadFromXml(QXmlStreamReader &reader);

private:
	friend class Worksheet;
	friend class WorksheetPrivate;
	QExplicitlySharedDataPointer<CellFormulaPrivate> d;
};

QT_END_NAMESPACE_XLSX
