#include <QDebug>

#include <QXlsx/xlsxcell.h>
#include <QXlsx/xlsxcelllocation.h>
#include <QXlsx/xlsxdocument.h>
#include <QXlsx/xlsxworkbook.h>
#include <QXlsx/xlsxworksheet.h>

void readXlsxDocument(const QString &filePath);

int main(int argc, char *argv[])
{
	if (argc < 2) {
		qDebug() << "Input file path of xlsx document";
	}

	readXlsxDocument(argv[1]);
	return 0;
}

void readXlsxDocument(const QString &filePath)
{
	qDebug() << filePath;

	QXlsx::Document doc(filePath);
	if (!doc.load()) {
		qDebug() << "Can't open file";
		return;
	}

	foreach (const auto &sheetName, doc.sheetNames()) {
		qDebug() << sheetName;
		const auto worksheet = static_cast<QXlsx::Worksheet *>(doc.sheet(sheetName));
		const auto cells = worksheet->getCells();
		for (const auto &c : cells) {
			qDebug() << "{ " << c.row << ", " << c.col << ", " << c.cell->readValue().toString() << " }";
		}
	}
}
