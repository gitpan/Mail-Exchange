package Mail::Exchange::PropertyContainer;

use strict;
use warnings;
use 5.008;

use Exporter;
use Encode;
use Mail::Exchange::PidTagIDs;
use Mail::Exchange::PidTagDefs;
use OLE::Storage_Lite;

use vars qw($VERSION @ISA);
@ISA=qw(Exporter);

$VERSION = "0.01";

sub new {
	my $class=shift;
	my $file=shift;

	my $self={};
	bless($self, $class);

	$self->{_properties}={};
	$self;
}

sub set {
	my $self=shift;
	my $property=shift;
	my $value=shift;
	my $flags=shift;
	my $type=shift;
	my $guid=shift;
	my $namedproperties = shift;

	my $normalized=$self->_propertyid($property, $type, $guid, $namedproperties);
	$self->{_properties}{$normalized}{val} = $value;
	$self->{_properties}{$normalized}{flg} = $flags;
}

sub get {
	my $self=shift;
	my $property=shift;
	my $namedproperties = shift;

	my $normalized=$self->_propertyid($property, undef, undef, $namedproperties);
	if (wantarray) {
		return ($self->{_properties}{$normalized}{val}, $self->{_properties}{$normalized}{flg});
	} else {
		return $self->{_properties}{$normalized}{val};
	}
}

sub _OlePropertyStreamlist {
	my $self=shift;
	my $unicode=shift;
	my $header=shift;

        my @streams=();
        my $propertystr=$header;

        foreach my $property(sort {$a <=> $b } keys %{$self->{_properties}}) {
		my $type;
                if ($property & 0x8000) {
                        $type=$self->{_namedProperties}->getType($property);
                } else {
                        $type=$PidTagDefs{$property}{type};
                }
                die "no type for $property" unless $type;
                # my $data=$self->get($property);
		my $data=$self->{_properties}{$property}{val};
		my $flags=$self->{_properties}{$property}{flg} || 6;

                if ($type==0x000d || $type==0x001e || $type==0x001f
                ||  $type==0x0048 || $type==0x0102) {
			no warnings 'portable';
			my $length;
                        if (($type == 0x001E || $type == 0x001F) && $unicode) {
                                $data=Encode::encode("UCS2LE", $data);
				$type=0x001F;
				$length=(length($data)+2) | 0x300000000;
                        } elsif (($type == 0x001E || $type == 0x001F) && !$unicode) {
                                $data=OLE::Storage_Lite::Ucs2Asc($data);
				$type=0x001E;
				$length=(length($data)+1) | 0x300000000;
                        } else {
				$length=(length($data)) | 0x300000000;
			}
                        my $streamname=sprintf("__substg1.0_%04X%04X", $property, $type);
                        my $stream=OLE::Storage_Lite::PPS::File->
                                new(Encode::encode("UCS2LE", $streamname), $data);
                        push(@streams, $stream);
                        $data=$length;
                }
                $propertystr.=pack("VVQ", ($property<<16|$type), $flags, $data);
        }
        my $stream=OLE::Storage_Lite::PPS::File->
                new(Encode::encode("UCS2LE", "__properties_version1.0"), $propertystr);
        push(@streams, $stream);

	return @streams;
}


# returns the internal hash index ID of a property,
# which is the upper 2 bytes of the official ID, without the
# lower 2 bytes that encode the type

sub _propertyid {
	my $self=shift;
	my $property=shift;
	my $type=shift;
	my $guid=shift;
	my $namedProperties=shift;

	if (substr($property, 0, 6) eq "PidTag") {
		die "Pid Tag decoding not implemented";
		# my $hash=$properties{$property};
		# die "$property unknown"
		#	unless $hash && $hash->{id};
		# return $hash->{id};
	} elsif (substr($property, 0, 6) eq "PidLid") {
		die "LID decoding not implemented";
		# get info, get guid, add to named properties
	} elsif ($property =~/^[0-9]/ && (($property&0xffff0000) == 0)
	    &&  ($property & 0x8000 || $PidTagDefs{$property})) {
		return $property;
	} elsif ($property =~/^[0-9]/ && $property & 0xffff0000) {
		my $id=($property>>16)&0xffff;
		return $id if $id & 0x8000;
		my $type=$property&0xffff;
		$type=0x1f if $type==0x1e;	# map String8 to UCS-String

		# When parsing, we might encounter an undocumented
		# property. Remember the type and hope for the best.
		if (!$PidTagDefs{$id}) {
			$PidTagDefs{$id}{type}=$type;
		}
		if ($PidTagDefs{$id}{type}!=$type) {
			die(sprintf("wrong type %04x for property %04x", $type, $id));
		}
		return $id;
	} elsif ($namedProperties) {
		# @@@ map guid name to guid ID ?
		my $id=$namedProperties->namedPropertyID($property, $type, $guid);
		return $id;
	}
	die ("can't make sense of $property");
}

sub _parseProperties {
	my $self=shift;
	my $file=shift;
	my $dir=shift;
	my $headersize=shift;
	my $namedProperties=shift;

	my $data=substr($file->{Data}, $headersize);	# ignore header
	while ($data) {
		my ($tag, $flags, $value)=unpack("VVQ", $data);
		my $type = $tag&0xffff;
		my $ptag = ($tag>>16)&0xffff;

		# Named property types aren't stored in __nameid so we have to set them now.
		if ($ptag & 0x8000) {
			$namedProperties->setType($ptag, $type);
		}
		if ($type & 0x1000) {
			die("Multiple properties not implemented");
		}
		if ($type==0x0002) { $value&=0xffff; }
		if ($type==0x0003 || $type==0x0004 || $type==0x000a || $type==0x000b
		||  $type==0x000d || $type==0x001e || $type==0x001f || $type==0x0048
		||  $type==0x00FB || $type==0x00FD || $type==0x00FE || $type==0x0102) {
			$value&=0xffffffff;
		}
		if ($type==0x000d || $type==0x001E || $type==0x001F
		||  $type==0x0048 || $type==0x0102) {
			my $streamname=Encode::encode("UCS2LE",
				sprintf("__substg1.0_%08X", $tag));
			my $found=0;
			foreach $file (@{$dir->{Child}}) {
				if ($file->{Name} eq $streamname) {
					$found=1;
					$value=$file->{Data};
					if ($type == 0x1f) {
						$value=Encode::decode("UCS2LE", $value);
					}
					last;
				}
			}
			die "stream for $tag not found" unless $found;
		}
		$self->set($tag, $value, $flags);
		$data=substr($data, 16);
	}
}
