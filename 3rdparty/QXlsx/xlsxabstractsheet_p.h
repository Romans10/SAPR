#pragma once

#include <QSharedPointer>
#include <QString>
#include <QtGlobal>

#include "xlsxglobal.h"

#include "xlsxabstractsheet.h"
#include "xlsxabstractxmlfile_p.h"
#include "xlsxdrawing_p.h"

QT_BEGIN_NAMESPACE_XLSX

class AbstractSheetPrivate : public AbstractXmlFilePrivate
{
	Q_DECLARE_PUBLIC(AbstractSheet)
public:
	AbstractSheetPrivate(AbstractSheet *p, AbstractSheet::CreateFlag flag);
	~AbstractSheetPrivate();

	Workbook *workbook;
	QSharedPointer<Drawing> drawing;

	QString name;
	int id;
	AbstractSheet::State sheetState;
	AbstractSheet::Type type;
};

QT_END_NAMESPACE_XLSX
