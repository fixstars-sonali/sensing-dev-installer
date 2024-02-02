// This checks if ion-kit installed under SENSING_DEV_ROOT directory.

#include <ion/ion.h>

int main(int argc, char *argv[])
{
    ion::Builder b;
    b.set_target(Halide::get_host_target());
    b.with_bb_module("ion-bb");
    return 0;
}