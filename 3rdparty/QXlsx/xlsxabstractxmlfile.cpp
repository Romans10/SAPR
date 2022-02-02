// xlsxabstractooxmlfile.cpp

#include <QBuffer>
#include <QByteArray>
#include <QtGlobal>

#include "xlsxabstractxmlfile.h"
#include "xlsxabstractxmlfile_p.h"

QT_BEGIN_NAMESPACE_XLSX

AbstractXmlFilePrivate::AbstractXmlFilePrivate(AbstractXmlFile *q, AbstractXmlFile::CreateFlag flag = AbstractXmlFile::F_NewFromScratch)
    : relationships(new Relationships)
    , flag(flag)
    , q_ptr(q)
{
}

AbstractXmlFilePrivate::~AbstractXmlFilePrivate()
{
}

/*!
 * \internal
 *
 * \class AbstractXmlFile
 *
 * Base class of all the ooxml part file.
 */

AbstractXmlFile::AbstractXmlFile(CreateFlag flag)
    : d_ptr(new AbstractXmlFilePrivate(this, flag))
{
}

AbstractXmlFile::AbstractXmlFile(AbstractXmlFilePrivate *d)
    : d_ptr(d)
{
}

AbstractXmlFile::~AbstractXmlFile()
{
	if (d_ptr->relationships)
		delete d_ptr->relationships;
	delete d_ptr;
}

QByteArray AbstractXmlFile::saveToXmlData() const
{
	QByteArray data;
	QBuffer buffer(&data);
	buffer.open(QIODevice::WriteOnly);
	saveToXmlFile(&buffer);

	return data;
}

bool AbstractXmlFile::loadFromXmlData(const QByteArray &data)
{
	QBuffer buffer;
	buffer.setData(data);
	buffer.open(QIODevice::ReadOnly);

	return loadFromXmlFile(&buffer);
}

/*!
 * \internal
 */
void AbstractXmlFile::setFilePath(const QString path)
{
	Q_D(AbstractXmlFile);
	d->filePathInPackage = path;
}

/*!
 * \internal
 */
QString AbstractXmlFile::filePath() const
{
	Q_D(const AbstractXmlFile);
	return d->filePathInPackage;
}

/*!
 * \internal
 */
Relationships *AbstractXmlFile::relationships() const
{
	Q_D(const AbstractXmlFile);
	return d->relationships;
}

QT_END_NAMESPACE_XLSX
