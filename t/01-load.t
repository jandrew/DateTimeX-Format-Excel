#!perl
### Test that the module(s) load!(s)
use	Test::More tests => 15;
use	lib '../lib', 'lib';
BEGIN{ use_ok( TAP::Harness ) };
BEGIN{ use_ok( Test::More ) };
BEGIN{ use_ok( Test::Moose ) };
BEGIN{ use_ok( Capture::Tiny ) };
BEGIN{ use_ok( version ) };
BEGIN{ use_ok(	Moose, 2.1213 ) };
BEGIN{ use_ok(	MooseX::StrictConstructor ) };
BEGIN{ use_ok(	MooseX::HasDefaults::RO ) };
BEGIN{ use_ok(	DateTime ) };
BEGIN{ use_ok(	Carp, qw( cluck ) ) };
BEGIN{ $ENV{PERL_TYPE_TINY_XS} = 0; };
BEGIN{ use_ok(	Type::Utils ) };
BEGIN{ use_ok(	Type::Library ) };
BEGIN{ use_ok(	Types::Standard ) };
BEGIN{ use_ok( DateTimeX::Format::Excel::Types, 0.014 ) };
BEGIN{ use_ok( DateTimeX::Format::Excel, 0.014 ) };
done_testing();
