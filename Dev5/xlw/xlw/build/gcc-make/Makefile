
ifndef BUILD
BUILD=DEBUG
endif

all:
	$(MAKE) -f InterfaceGenerator.mak   BUILD=$(BUILD)
	$(MAKE) -f XLW.mak BUILD=$(BUILD)
	
clean:
	$(MAKE) -f InterfaceGenerator.mak  BUILD=$(BUILD) clean
	$(MAKE) -f XLW.mak BUILD=$(BUILD) clean
	
rebuild: clean all

