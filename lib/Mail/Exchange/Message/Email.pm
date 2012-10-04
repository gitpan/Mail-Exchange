package Mail::Exchange::Message::Email;

use Mail::Exchange::PidTagIDs;
use Mail::Exchange::Message;
use Mail::Exchange::Recipient;

=head1 NAME

Mail::Exchange::Message::Email - subclass of Mail::Exchange::Message
that initializes Email-specific fields

=head1 SYNOPSIS

    use Mail::Exchange::Message::Email;

    $mail=Mail::Exchange::Message::Email->new();

=head1 DESCRIPTION

Mail::Exchange::Message::Email is a utility class derived from
Mail::Exchange::Message. When creating a new message object, it sets the
Message Class to "IPM.Note" to mark this message as an email object.

=cut

use strict;
use warnings;
use 5.008;

use Exporter;

use vars qw($VERSION @ISA);
@ISA=qw(Mail::Exchange::Message Exporter);

$VERSION="0.02";

=head2 new()

$msg=Mail::Exchange::Message::Email->new();

Create a new message object and initialize it to an email.
=cut

sub new {
	my $class=shift;
	my $self=Mail::Exchange::Message->new();
	$self->set(PidTagMessageClass, "IPM.Note");
	bless $self;
}

=head2 parse()

The parse() method is overwritten to abort, because the message type will be
read from the input file, so a plain Mail::Exchange::Message object should
be used in this case.

=cut

sub parse {
	die("parse not supported, use a Mail::Exchange::Message object");
}

=head2 fromMIME()

fromMIME takes an Email::MIME object and returns a 
Mail::Exchange::Message::Email object.

=cut

sub fromMIME($) {
	my $class=shift;
	my $mime=shift;

	my $self=$class->new;

	die "wrong object type" if ref $mime ne "Email::MIME";

	$self->set(PidTagStoreSupportMask, 0x40000);
	$self->set(PidTagDisplayTo, join(", ", $mime->header("To")));
	$self->set(PidTagDisplayCc, join(", ", $mime->header("Cc")));
	$self->set(PidTagDisplayBcc, join(", ", $mime->header("Bcc")));
	$self->set(PidTagInternetMessageId, $mime->header("Message-ID"));
	$self->set(PidTagMessageFlags, 1);
	$self->set(PidTagSenderAddressType, "SMTP");
	$self->set(PidTagSenderEmailAddress, $mime->header("From"));
	$self->set(PidTagSenderName, $mime->header("From"));
	$self->set(PidTagSenderSmtpAddress, $mime->header("From"));
	$self->setSubject($mime->header("Subject"));
	$self->set(PidTagTransportMessageHeaders,
		$mime->header_obj->as_string());

	# unfortunately, Email::MIME::walk_parts can' pass anything
	# through to the callback function so we can't use it.

	my @parts=($mime);
	while (my $part=pop @parts) {
		push(@parts, $part->subparts);
		next if ($part->{ct}{discrete} eq "multipart"
		     ||  $part->{ct}{discrete} eq "message");

		# If it has a filename, assume it's an attachment.
		# If it doesn't, and it's text/plain or text/html, set the
		# appropriate body part. Else invent a filename and attach it.
		my $filename;
		my $attach;
		if ($filename=$part->filename()) {
			$attach=Mail::Exchange::Attachment->new();
			$attach->setFileName($filename);
		} elsif ($part->{ct}{discrete} eq "text"
		 &&   ($part->{ct}{composite} eq 'plain')) {
		 	$self->setBody($part->body);
			next;
		} elsif ($part->{ct}{discrete} eq "text"
		 &&   ($part->{ct}{composite} eq 'html')) {
		 	$self->setHTMLBody($part->body);
			next;
		} else {
			$attach=Mail::Exchange::Attachment->new();
			$attach->setFileName($part->invent_filename(
					$part->header("Content-type")));
		}
		$attach->setString($part->body);
		if (my $cid=$part->header("Content-ID")) {
			$attach->set(PidTagAttachContentId, $cid);
		}
		$self->addAttachment($attach);
	}

	foreach my $type qw(To Cc Bcc) { if ($mime->header($type)) {
		foreach my $recipname (split(",", $mime->header($type))) {
			my $recipient=Mail::Exchange::Recipient->new();
			$recipient->setRecipientType($type);
			$recipient->setDisplayName($recipname);
			$recipient->setEmailAddress($recipname);
			$self->addRecipient($recipient);
		}
	}}

	$self;
}
