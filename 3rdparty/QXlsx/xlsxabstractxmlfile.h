#pragma once

#include <QIODevice>

#include "xlsxglobal.h"

QT_BEGIN_NAMESPACE_XLSX

class Relationships;
class AbstractXmlFilePrivate;

class AbstractXmlFile
{
	Q_DECLARE_PRIVATE(AbstractXmlFile)

public:
	enum CreateFlag
	{
		F_NewFromScratch,
		F_LoadFromExists
	};

public:
	virtual ~AbstractXmlFile();

	virtual void saveToXmlFile(QIODevice *device) const = 0;
	virtual bool loadFromXmlFile(QIODevice *device) = 0;

	virtual QByteArray saveToXmlData() const;
	virtual bool loadFromXmlData(const QByteArray &data);

	Relationships *relationships() const;

	void setFilePath(const QString path);
	QString filePath() const;

protected:
	AbstractXmlFile(CreateFlag flag);
	AbstractXmlFile(AbstractXmlFilePrivate *d);

	AbstractXmlFilePrivate *d_ptr;
};

QT_END_NAMESPACE_XLSX
