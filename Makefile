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

oxt:
	rm -f *.oxt
	zip -r $(OXT) $(FILES)

commit:
	rm -f *.oxt
	git add ./* && git commit && git push
