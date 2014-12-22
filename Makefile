run_testing: testing
	killall soffice.bin || echo "No libreoffice instance found"
	lowriter "$$TESTING_ODT" --norestore &

testing: src/vibreoffice.vbs
	./compile.sh "src/vibreoffice.vbs" "$$TESTING_XBA"

extension: clean src/vibreoffice.vbs
	if [ -z "$$VIBREOFFICE_VERSION" ]; then \
		echo "VIBREOFFICE_VERSION must be set"; \
	else \
		mkdir -p build; mkdir -p dist; \
		cp -r extension/template build/template; \
		./compile.sh "src/vibreoffice.vbs" "build/template/vibreoffice/vibreoffice.xba"; \
		cd "build/template"; \
		sed -i "s/%VIBREOFFICE_VERSION%/$$VIBREOFFICE_VERSION/g" description.xml; \
		zip -r "../../dist/vibreoffice-$$VIBREOFFICE_VERSION.oxt" .; \
	fi

.PHONY: clean
clean:
	rm -rf build
