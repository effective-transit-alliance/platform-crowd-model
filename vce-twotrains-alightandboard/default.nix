with import <nixpkgs> {};

rec {
  pythonEnv = python3.buildEnv.override {
    extraLibs = with python3Packages; [
      black
      flake8
      openpyxl
      numpy
    ];
  };

  build = runCommand "platform_F_twotrains_twoway_new.xlsx" {
	nativeBuildInputs = [ pythonEnv ];
  } ''
    python ${./vce-twotrains-alightandboard.py}
    mv *.xlsx "$out"
  '';

  shell = pythonEnv.env;
}
