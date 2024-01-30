// This checks if aravis and gobject are installed under SENSING_DEV_ROOT directory.

#include <exception>
#include <iostream>
#include "arv.h"

int main(int argc, char *argv[])
{
    arv_update_device_list ();
    return 0;
}