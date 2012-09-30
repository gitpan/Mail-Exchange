package Mail::Exchange::Recipient;

use strict;
use warnings;
use 5.008;

use Exporter;
use Encode;
use Mail::Exchange::PidTagDefs;
use Mail::Exchange::PidTagIDs;
use Mail::Exchange::PropertyContainer;

use vars qw($VERSION @ISA);
@ISA=qw(Mail::Exchange::PropertyContainer Exporter);

$VERSION = "0.01";

sub new {
	my $class=shift;
	my $file=shift;

	my $self=Mail::Exchange::PropertyContainer->new();
	bless($self, $class);

	$self->set(PidTagRowid, 1);
	$self->set(PidTagRecipientType, 1);
	$self->set(PidTagDisplayType, 0);
	$self->set(PidTagObjectType, 6);
	$self->set(PidTagAddressType, "SMTP");

	$self;
}

sub setRecipientType {
	my $self=shift;
	my $field=shift;

	my $type=0;
	if (uc $field eq "TO")	{ $type=1; }
	if (uc $field eq "CC")	{ $type=2; }
	if (uc $field eq "BCC")	{ $type=3; }
	if ($field =~ /^[0-9]+$/) { $type=$field; }

	die "unknown Recipient Type $field" if ($type==0);

	$self->set(PidTagRecipientType, $type);
}

sub setAddressType {
	my $self=shift;
	my $type=shift;

	$self->set(PidTagAddressType, $type);
}

sub setDisplayName {
	my $self=shift;
	my $name=shift;

	$self->set(PidTagDisplayName, $name);
	$self->set(PidTagTransmittableDisplayName, $name);
}

sub setSMTPAddress {
	my $self=shift;
	my $recipient=shift;

	$self->set(PidTagSmtpAddress, $recipient);
}

sub setEmailAddress {
	my $self=shift;
	my $recipient=shift;

	$self->set(PidTagSmtpAddress, $recipient);
	$self->set(PidTagEmailAddress, $recipient);
}

sub OleContainer {
	my $self=shift;
	my $no=shift;
	my $unicode=shift;

	my $header=pack("V2", 0, 0);

	$self->set(PidTagRowid, $no);

	my @streams=$self->_OlePropertyStreamlist($unicode, $header);
	my $dirname=Encode::encode("UCS2LE", sprintf("__recip_version1.0_#%08X", $no));
	my @ltime=localtime();
	my $dir=OLE::Storage_Lite::PPS::Dir->new($dirname, \@ltime, \@ltime, \@streams);
	return $dir;
}

sub _parseRecipientProperties {
	my $self=shift;
	my $file=shift;
	my $dir=shift;
	my $namedProperties=shift;

	$self->_parseProperties($file, $dir, 8, $namedProperties);
}

1;
