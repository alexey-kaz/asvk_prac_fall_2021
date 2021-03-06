#include <com/sun/star/lang/XMultiServiceFactory.hpp>

#include <com/sun/star/frame/XComponentLoader.hpp>

#include <com/sun/star/text/XTextDocument.hpp>

#include <com/sun/star/text/XText.hpp>

#include <com/sun/star/text/XTextTable.hpp>

#include <com/sun/star/text/XTextContent.hpp>

#include <com/sun/star/table/XCell.hpp>

#include <com/sun/star/table/XTable.hpp>

#include "PluginFunc.h"

#include <map>

#include <vector>


using namespace com::sun::star::uno;
using namespace com::sun::star::lang;
using namespace com::sun::star::beans;
using namespace com::sun::star::frame;
using namespace com::sun::star::text;
using namespace com::sun::star::table;
using namespace com::sun::star::util;
using ::rtl::OUString;

std::u16string
const cyrillic = u"АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдежзийклмнопрстуфхцчшщъыьэюя";
std::u16string
const latin = u"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
std::u16string
const mix = cyrillic + latin;

std::u16string langs[3] = {cyrillic, latin, mix};

void New_Doc_Creator(Reference <XFrame> &rxFrame, int global_lang, int globalWordCnt, int global_word_len) {
    if (not rxFrame.is())
        return;

    Reference <XComponentLoader> rComponentLoader(rxFrame, UNO_QUERY);
    Reference <XComponent> xWriterComponent = rComponentLoader->loadComponentFromURL(
            OUString::createFromAscii("private:factory/swriter"),
            OUString::createFromAscii("_blank"),
            0,
            Sequence<::com::sun::star::beans::PropertyValue>());
    Reference <XTextDocument> xTextDocument(xWriterComponent, UNO_QUERY);
    Reference <XText> xText = xTextDocument->getText();
    Reference <XTextCursor> xTextCursor = xText->createTextCursor();
    OUString text = "";
    srand(time(NULL));
    for (std::size_t i = 0; i < globalWordCnt; i++) {
        OUString word = "";
        for (std::size_t j = 0; j < (1 + std::rand() % (global_word_len)); j++) {
            word += (OUString) langs[global_lang - 1][rand() % langs[global_lang - 1].length()];
        }
        text += word + " ";

    }
    xText->insertString(xTextCursor, text, false);
    return;
}

std::multimap<int, sal_Unicode> sort_mm(std::map<sal_Unicode, int>& m)
{
    std::multimap<int, sal_Unicode> multi;
    for (auto& item: m){
        multi.insert({item.second, item.first});
    }
    return multi;
}

void Stats(Reference <XFrame> &mxFrame) {
    std::map<sal_Unicode, int> stat_Table;
    Reference <XTextDocument> xTextDocument(mxFrame->getController()->getModel(), UNO_QUERY);
    Reference <XText> xText = xTextDocument->getText();
    Reference <XTextCursor> xTextCursor = xText->createTextCursor();
    Reference <XPropertySet> xCursorProps(xTextCursor, UNO_QUERY);

    while (xTextCursor->goRight(1, true)) {
        char16_t chr = (xTextCursor->getString()).toChar();
        if (mix.find(chr) != -1)
            stat_Table[chr]++;
        xTextCursor->collapseToEnd();
    }

    xTextCursor->gotoEnd(false);

    Reference <XMultiServiceFactory> oDocMSF(xTextDocument, UNO_QUERY);
    Reference <XTextTable> xTable(oDocMSF->createInstance(
            OUString::createFromAscii("com.sun.star.text.TextTable")), UNO_QUERY);

    xTable->initialize(stat_Table.size() + 1, 2);
    Reference <XTextRange> xTextRange = xText->getEnd();
    Reference <XTextContent> xTextContent(xTable, UNO_QUERY);
    xText->insertTextContent(xTextRange, xTextContent, (unsigned char) 0);

    Reference <XCell> xCell;
    int i = 0;
    std::array<std::pair<const char *, const char *>, 2> conf = {std::make_pair("A1", "Char"),
                                                                 std::make_pair("B1", "Char Count")};
    for (i; i < 2; i++) {
        xCell = xTable->getCellByName(OUString::createFromAscii(conf[i].first));
        xText = Reference<XText>(xCell, UNO_QUERY);
        xTextCursor = xText->createTextCursor();
        xTextCursor->setString(OUString::createFromAscii(conf[i].second));
    }
    std::multimap<int, sal_Unicode> stat_Table_mm = sort_mm(stat_Table);
    for (std::multimap<int, sal_Unicode>::reverse_iterator it = stat_Table_mm.rbegin(); it != stat_Table_mm.rend(); ++it, i++) {
        xCell = xTable->getCellByName(OUString::createFromAscii(((char) ('A') + std::to_string(i)).c_str()));
        xTextCursor = Reference<XText>(xCell, UNO_QUERY)->createTextCursor();
        xTextCursor->setString(OUString(it->second));
        xCell = xTable->getCellByName(OUString::createFromAscii(((char) ('B') + std::to_string(i)).c_str()));
        xCell->setValue(it->first);
    }
    return;

}

bool withCyrillicLetters(rtl::OUString str) {
    for (int i = 0; i < str.getLength(); i++) {
        for (auto chr: cyrillic)
            if (str[i] == chr) {
                return true;
            }
    }
    return false;
}

bool Is_Not_Special(char16_t chr) {
    return mix.find(chr) != -1 || std::isdigit(chr);
}

void Highlight_Cyrillic(Reference <XFrame> &mxFrame) {
    Reference <XTextDocument> xTextDocument(mxFrame->getController()->getModel(), UNO_QUERY);
    Reference <XText> xText = xTextDocument->getText();
    Reference <XTextCursor> xTextCursor = xText->createTextCursor();
    std::size_t curSum = 0;
    std::size_t txtLen = (xText->getString()).getLength();
    Reference <XPropertySet> xCursorProps(xTextCursor, UNO_QUERY);
    while (xTextCursor->goRight(1, true)) {
        std::size_t curSize = 0;
        while (curSize < (txtLen - curSum) && Is_Not_Special((xTextCursor->getString())[curSize])) {
            xTextCursor->goRight(1, true);
            curSize++;
        }
        if (curSize != (txtLen - curSum)) {
            xTextCursor->goLeft(1, true);
        }
        if (withCyrillicLetters(xTextCursor->getString())) {
            xCursorProps->setPropertyValue("CharColor", makeAny(0xFF0000));
        }
        if (xTextCursor->goRight(1, true)) {
            xTextCursor->collapseToEnd();
            curSum += curSize + 1;
        } else {
            break;
        }
    }
    return;
}