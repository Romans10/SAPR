#pragma once

#include "xlsxabstractxmlfile.h"
#include "xlsxglobal.h"

QT_BEGIN_NAMESPACE_XLSX

class Workbook;
class Drawing;
class AbstractSheetPrivate;

class AbstractSheet : public AbstractXmlFile
{
	Q_DECLARE_PRIVATE(AbstractSheet)

public:
	Workbook *workbook() const;

	enum class Type
	{
		Work,
		Chart,
		Dialog,
		Macro
	};

	enum class State
	{
		Visible,
		Hidden,
		VeryHidden
	};

	QString sheetName() const;
	Type sheetType() const;
	State sheetState() const;
	void setState(State ss);
	bool isHidden() const;
	bool isVisible() const;
	void setHidden(bool hidden);
	void setVisible(bool visible);

protected:
	friend class Workbook;
	AbstractSheet(const QString &sheetName, int sheetId, Workbook *book, AbstractSheetPrivate *d);
	virtual AbstractSheet *copy(const QString &distName, int distId) const = 0;
	void setSheetName(const QString &sheetName);
	void setSheetType(Type type);
	int sheetId() const;

	Drawing *drawing() const;
};

QT_END_NAMESPACE_XLSX
