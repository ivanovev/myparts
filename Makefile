FILES=MP.py \
	pythonpath/mp/__init__.py \
	pythonpath/mp/myparts.py \
	pythonpath/mp/mysort.py \
	pythonpath/mp/mybom.py \
	Addons.xcu \
	description.xml \
	Office/UI/CalcWindowState.xcu \
	images/Logo.png \
	images/part_16.bmp \
	images/part_26.bmp \
	images/sort_16.bmp \
	images/sort_26.bmp \
	images/bom_16.bmp \
	images/bom_26.bmp \
	META-INF/manifest.xml \
	pkg-desc/pkg-description.txt
VER=$(shell date +%F)
OXT=myparts-$(VER).oxt

.PHONY: update_version commit

update_version:
	cat description.xml | sed "s:  <version value=[^^]* />:  <version value=\"$(shell date +%F)\" />:" > desc.xml
	mv desc.xml description.xml

oxt: update_version
	rm -f *.oxt
	zip -r $(OXT) $(FILES)

commit:
	rm -f *.oxt
	git add ./* && git commit && git push
