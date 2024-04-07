with import <nixpkgs> {};

(python3.buildEnv.override {
  extraLibs = with python3Packages; [
    black
    flake8
    openpyxl
    numpy
  ];
}).env
