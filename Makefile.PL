use ExtUtils::MakeMaker;

$VERSION="0.03.1";

WriteMakefile(
    ($] >= 5.008 ? (
    'AUTHOR'        => 'Guntram Blohm (gbl@bso2001.com)',
    'ABSTRACT'      => 'Read and write Outlook .msg files', ) : () ),
    'NAME'          => 'Mail::Exchange',
    'VERSION_FROM'  => 'Makefile.PL',
    'NEEDS_LINKING' => 0,
    'PREREQ_PM'     => {
    		'Email::Address' => "1.892",
    		'OLE::Storage_Lite' => "0.19",
    },
    'dist'          => {COMPRESS => 'gzip --best', SUFFIX => 'gz'},


);
