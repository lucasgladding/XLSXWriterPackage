# XLSXWriterPackage

This package provides a Swift wrapper around libxlsxwriter. You can find more information about libxlsxwriter here:
https://github.com/jmcnamara/libxlsxwriter

I expect to add some docs to this package at some point, but in the meantime please check the Tests folder for some examples. The package was developed by taking the examples from the library and rewriting them using an interface that I wanted to work with. The "tests" do not contain assertions, but generate files that can be reviewed for issues. I am not logging the filepath, but you can add a breakpoint in the tests and inspect the path from there. The examples that serve as the basis can be found here:
https://libxlsxwriter.github.io/examples.html
