#pragma once

#include "xlsxabstractxmlfile.h"
#include "xlsxglobal.h"
#include "xlsxrelationships_p.h"

QT_BEGIN_NAMESPACE_XLSX

class AbstractXmlFilePrivate
{
	Q_DECLARE_PUBLIC(AbstractXmlFile)

public:
	AbstractXmlFilePrivate(AbstractXmlFile *q, AbstractXmlFile::CreateFlag flag);
	virtual ~AbstractXmlFilePrivate();

	QString filePathInPackage;  // such as "xl/worksheets/sheet1.xml"

	Relationships *relationships;
	AbstractXmlFile::CreateFlag flag;
	AbstractXmlFile *q_ptr;
};

QT_END_NAMESPACE_XLSX
