#pragma once

#include <QList>
#include <QSharedPointer>
#include <QString>
#include <QtGlobal>

#include "xlsxabstractxmlfile.h"
#include "xlsxrelationships_p.h"

class QIODevice;
class QXmlStreamWriter;

QT_BEGIN_NAMESPACE_XLSX

class DrawingAnchor;
class Workbook;
class AbstractSheet;
class MediaFile;

class Drawing : public AbstractXmlFile
{
public:
	Drawing(AbstractSheet *sheet, CreateFlag flag);
	~Drawing();

	void saveToXmlFile(QIODevice *device) const;
	bool loadFromXmlFile(QIODevice *device);

	AbstractSheet *sheet;
	Workbook *workbook;
	QList<DrawingAnchor *> anchors;
};

QT_END_NAMESPACE_XLSX
