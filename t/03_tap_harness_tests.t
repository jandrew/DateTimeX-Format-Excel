#!perl
#~ use Smart::Comments '###';
my	$dir 	= './';
my	$up		= '../';
for my $next ( <*> ){
	if( ($next eq 't') and -d $next ){
		### <where> - found the t directory - must be using prove ...
		$dir	= './t/';
		$up		= '';
		last;
	}
}
### <where> - dir is: $dir
### <where> - up is: $up
		
my	$args ={
		#~ verbosity => 1,
		lib =>[
			$up . 'lib',
		],
		test_args =>{
			types_test 							=> '',
			date_time_original_basic_test		=> '',
			date_time_original_fractions_test	=> '',
			excel_module_test					=>[ $dir . 'test_files/' ],
		}
	};
### <where> - args: $args
my	@tests =(
		[  $dir . 'DateTimeX/Format/Excel/01_types.t', 	'types_test' ],
		[  $dir . 'DateTimeX/Format/00_excel.t', 		'excel_module_test' ],
		[  $dir . 'DateTimeX/Format/01_basic.t', 		'date_time_original_basic_test' ],
		[  $dir . 'DateTimeX/Format/02_fractions.t', 	'date_time_original_fractions_test' ],
	);
use	TAP::Harness;
my	$harness = TAP::Harness->new( $args );
	$harness->runtests(@tests);
use Test::More;
pass( "Finished the TAP Harness tests" );
done_testing();