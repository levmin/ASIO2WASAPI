#pragma once
#define VERSION_MAJOR 1
#define VERSION_MINOR 2
#define VERSION_PATCH 1
#define VERSION_BUILD 1

#define stringify(a) stringify_(a)
#define stringify_(a) #a

#define PRODUCT_VERSION VERSION_MAJOR, VERSION_MINOR, VERSION_PATCH
#define SPRODUCT_VERSION stringify(VERSION_MAJOR) "." stringify(VERSION_MINOR) "." stringify(VERSION_PATCH)
#define FILE_VERSION PRODUCT_VERSION, VERSION_BUILD
#define SFILE_VERSION SPRODUCT_VERSION "." stringify(VERSION_BUILD)
