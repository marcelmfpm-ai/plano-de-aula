{ pkgs }: {
  deps = [
    pkgs.python3
    pkgs.python3Packages.pip
    pkgs.stdenv.cc.cc.lib
    pkgs.zlib
    pkgs.gcc
  ];
  env = {
    LD_LIBRARY_PATH = pkgs.lib.makeLibraryPath [
      pkgs.stdenv.cc.cc.lib
      pkgs.zlib
    ];
  };
}
