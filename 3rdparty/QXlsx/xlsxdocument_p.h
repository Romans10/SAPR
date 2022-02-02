#pragma once

#include <QMap>
#include <QtGlobal>

#include "xlsxcontenttypes_p.h"
#include "xlsxdocument.h"
#include "xlsxglobal.h"
#include "xlsxworkbook.h"

QT_BEGIN_NAMESPACE_XLSX

class DocumentPrivate
{
	Q_DECLARE_PUBLIC(Document)
public:
	explicit DocumentPrivate(Document *p);
	void init();

	bool loadPackage(QIODevice *device);
	bool savePackage(QIODevice *device) const;

	// copy style from one xlsx file to other
	static bool copyStyle(const QString &from, const QString &to);

	Document *q_ptr;
	const QString defaultPackageName;  // default name when package name not specified
	QString packageName;  // name of the .xlsx file

	QMap<QString, QString> documentProperties;  // core, app and custom properties
	QSharedPointer<Workbook> workbook;
	QSharedPointer<ContentTypes> contentTypes;
	bool isLoad;
};

QT_END_NAMESPACE_XLSX
