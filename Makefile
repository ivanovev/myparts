FILES=myparts.py \
	Addons.xcu \
	description.xml \
	Office/UI/CalcWindowState.xcu \
	images/Logo.png \
	images/part_16.bmp \
	images/part_26.bmp \
	META-INF/manifest.xml
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
