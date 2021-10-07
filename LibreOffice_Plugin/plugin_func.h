using namespace com::sun::star::uno;
using namespace com::sun::star::frame;

void New_Doc_Creator(Reference< XFrame > &rxFrame, int global_lang,  int globalWordCnt, int global_word_len);
void Highlight_Cyrillic(Reference< XFrame > &rxFrame );
void Stats (Reference< XFrame > &mxFrame );
bool withCyrillicLetters(rtl::OUString curWord);
bool Is_Not_Special(char16_t letter);