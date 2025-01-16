CMD := wkhtmltopdf
CMDOPT := --enable-local-file-access -s A5
TODAY := `date "+%d %b %Y"`
FOOTOPT := --footer-left "[section] - [subsection]" --footer-right "[page]" --footer-font-size 8 
TOCOPT := toc --xsl-style-sheet toc.xsl
TOCFOOTOPT :=  --header-center "Published $(TODAY)" --header-font-size 6
OUTLINEOPT := --outline-depth 1

TARGETS := UserManual.pdf

UserManual.pdf:	UserManual.html
	-${CMD} ${CMDOPT} ${OUTLINEOPT} ${TOCOPT} ${TOCFOOTOPT}  UserManual.html ${FOOTOPT} UserManual.pdf