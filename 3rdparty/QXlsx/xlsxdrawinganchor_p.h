#pragma once

#include <QPoint>
#include <QSharedPointer>
#include <QSize>
#include <QString>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include "xlsxglobal.h"

QT_BEGIN_NAMESPACE_XLSX

class Drawing;
class MediaFile;
class Chart;

// Helper class
struct XlsxMarker
{
	XlsxMarker()
	{
	}

	XlsxMarker(int row, int column, int rowOffset, int colOffset)
	    : cell(QPoint(row, column))
	    , offset(rowOffset, colOffset)
	{
	}

	int row() const
	{
		return cell.x();
	}

	int col() const
	{
		return cell.y();
	}

	int rowOff() const
	{
		return offset.width();
	}

	int colOff() const
	{
		return offset.height();
	}

	QPoint cell;
	QSize offset;
};

class DrawingAnchor
{
public:
	enum class ObjectType
	{
		GraphicFrame,
		Shape,
		GroupShape,
		ConnectionShape,
		Picture,
		Unknown
	};

	DrawingAnchor(Drawing *drawing, ObjectType objectType);
	virtual ~DrawingAnchor();

	void setObjectPicture(const QImage &img);
	bool getObjectPicture(QImage &img);

	void setObjectGraphicFrame(QSharedPointer<Chart> chart);

	virtual bool loadFromXml(QXmlStreamReader &reader) = 0;
	virtual void saveToXml(QXmlStreamWriter &writer) const = 0;

	virtual int row();
	virtual int col();

	int getm_id();

protected:
	QPoint loadXmlPos(QXmlStreamReader &reader);
	QSize loadXmlExt(QXmlStreamReader &reader);
	XlsxMarker loadXmlMarker(QXmlStreamReader &reader, const QString &node);
	void loadXmlObject(QXmlStreamReader &reader);
	void loadXmlObjectShape(QXmlStreamReader &reader);
	void loadXmlObjectGroupShape(QXmlStreamReader &reader);
	void loadXmlObjectGraphicFrame(QXmlStreamReader &reader);
	void loadXmlObjectConnectionShape(QXmlStreamReader &reader);
	void loadXmlObjectPicture(QXmlStreamReader &reader);

	void saveXmlPos(QXmlStreamWriter &writer, const QPoint &pos) const;
	void saveXmlExt(QXmlStreamWriter &writer, const QSize &ext) const;
	void saveXmlMarker(QXmlStreamWriter &writer, const XlsxMarker &marker, const QString &node) const;
	void saveXmlObject(QXmlStreamWriter &writer) const;
	void saveXmlObjectShape(QXmlStreamWriter &writer) const;
	void saveXmlObjectGroupShape(QXmlStreamWriter &writer) const;
	void saveXmlObjectGraphicFrame(QXmlStreamWriter &writer) const;
	void saveXmlObjectConnectionShape(QXmlStreamWriter &writer) const;
	void saveXmlObjectPicture(QXmlStreamWriter &writer) const;

	Drawing *m_drawing;
	ObjectType m_objectType;
	QSharedPointer<MediaFile> m_pictureFile;
	QSharedPointer<Chart> m_chartFile;

	int m_id;

	// {{liufeij
	void setObjectShape(const QImage &img);  // liufeij

	QString editASName;
	// below only for twocellanchor shape
	QPoint posTA;  // for shape liufeij 20181024
	QSize extTA;  // for shape liufeij 20181024
	int rotWithShapeTA;  //// for shape liufeij 20181024
	int dpiTA;  //// for shape liufeij 20181024
	QString sp_textlink, sp_macro, sp_blip_cstate, sp_blip_rembed;

	// BELOW only for cxnSp shape
	QString cxnSp_filpV, cxnSp_macro;
	// belwo for cxnsp and sp
	QString xsp_cNvPR_name, xsp_cNvPR_id;  // x measns shape and cxnSp together using
	QString xbwMode;  // same as above
	QString xIn_algn, xIn_cmpd, xIn_cap, xIn_w;  // cxnSp only need xIn_w
	QString xprstGeom_prst;
	QString x_headEnd_w, x_headEnd_len, x_headEnd_tyep;
	QString x_tailEnd_w, x_tailEnd_len, x_tailEnd_tyep;
	QString Style_inref_idx, style_fillref_idx, style_effectref_idx, style_forntref_idx;
	QString Style_inref_val, style_fillref_val, style_effectref_val, style_forntref_val;
	// liufeij }}
};

class DrawingAbsoluteAnchor : public DrawingAnchor
{
public:
	DrawingAbsoluteAnchor(Drawing *drawing, ObjectType objectType = ObjectType::Unknown);

	QPoint pos;
	QSize ext;

	bool loadFromXml(QXmlStreamReader &reader);
	void saveToXml(QXmlStreamWriter &writer) const;
};

class DrawingOneCellAnchor : public DrawingAnchor
{
public:
	DrawingOneCellAnchor(Drawing *drawing, ObjectType objectType = ObjectType::Unknown);

	XlsxMarker from;
	QSize ext;

	virtual int row();
	virtual int col();

	bool loadFromXml(QXmlStreamReader &reader);
	void saveToXml(QXmlStreamWriter &writer) const;
};

class DrawingTwoCellAnchor : public DrawingAnchor
{
public:
	explicit DrawingTwoCellAnchor(Drawing *drawing, ObjectType objectType = ObjectType::Unknown);

	XlsxMarker from;
	XlsxMarker to;

	virtual int row();
	virtual int col();

	bool loadFromXml(QXmlStreamReader &reader);
	void saveToXml(QXmlStreamWriter &writer) const;
};

QT_END_NAMESPACE_XLSX
