CRUFT :=  foo bar
OTHER_TARGETS := 
TARGETS := GEDmatchImages

GEDmatchImages:
	mkdir GEDmatchImages
	cp */*.gif GEDmatchImages

rename:	${TARGETS} ${OTHER_TARGETS}
	cd GEDmatchImages &&  ls -1 *.gif > foo
	cd GEDmatchImages && awk -f ../renameImages.awk foo > bar
	cd GEDmatchImages && sh ./bar

mostlyclean:
	-rm -rf ${CRUFT}

clean:	mostlyclean	
	-rm -rf ${OTHER_TARGETS}

clobber:	clean
	-rm -rf ${TARGETS}

