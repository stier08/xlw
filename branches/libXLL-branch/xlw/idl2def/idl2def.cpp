
#include <iostream>

using namespace std;

int main(int argc, char* argv[])
{
	cout << "; generated by idl2def, do not edit!" << endl;
	cout << "; (c) 2001-2002 by Jens Thiel <Jens@Thiel.de>" << endl << endl;
	cout << "EXPORTS" << endl << endl;

	char buf[255];
	char *pdest;
	while( cin) {
		cin.getline(buf,256,'\n');
		if( pdest = strstr( buf, "XLL_EXPORT" ) ) {
			pdest += 11;
			pdest[strspn( pdest, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890_" )] = '\0';
			cout << "    " << pdest << endl;
		}
	}

	return 0;

}