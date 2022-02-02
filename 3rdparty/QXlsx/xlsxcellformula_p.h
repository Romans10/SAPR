#pragma once

#include "xlsxcellformula.h"
#include "xlsxcellrange.h"
#include "xlsxglobal.h"

#include <QSharedData>
#include <QString>

QT_BEGIN_NAMESPACE_XLSX

class CellFormulaPrivate : public QSharedData
{
public:
	CellFormulaPrivate(const QString &formula, const CellRange &reference, CellFormula::Type type);
	CellFormulaPrivate(const CellFormulaPrivate &other);
	~CellFormulaPrivate();

	QString formula;  // formula contents
	CellFormula::Type type;
	CellRange reference;
	bool ca;  // Calculate Cell
	int si;  // Shared group index
};

QT_END_NAMESPACE_XLSX
