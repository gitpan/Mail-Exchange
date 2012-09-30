use ExtUtils::MakeMaker;


WriteMakefile(
    ($] >= 5.008 ? (
    'AUTHOR'        => 'Guntram Blohm (gbl@bso2001.com)',
    'ABSTRACT'      => 'Read and write Outlook .msg files', ) : () ),
    'NAME'          => 'Mail::Exchange',
    'VERSION_FROM'  => 'lib/Mail/Exchange/Message.pm',
    'NEEDS_LINKING' => 0,
    'PREREQ_PM'     => {'OLE::Storage_Lite' => "0.19"},
    'dist'          => {COMPRESS => 'gzip --best', SUFFIX => 'gz'},


);
