<?php

ini_set('display_errors', '1');

require 'vendor/autoload.php';

use Hfig\MAPI;
use Hfig\MAPI\OLE\Pear;

// message parsing and file IO are kept separate
$messageFactory = new MAPI\MapiMessageFactory();
$documentFactory = new Pear\DocumentFactory(); 
$parser = new Phemail\MessageParser();

class extractSender {
    public function parseMsg($file) {
        $messageFactory = new MAPI\MapiMessageFactory();
        $documentFactory = new Pear\DocumentFactory(); 

        $ole = $documentFactory->createFromFile($file);
        $message = $messageFactory->parseMessage($ole);

        return $message->properties['sender_email_address'];
    }

    public function parseEml($file) {
        $parser = new Phemail\MessageParser();

        $message = $parser->parse($file);
        preg_match('/<([^#]+)>/', $message->getHeaderValue('from'), $match);
        return $match[1];
    }
}

class fileHandler {
    private function checkExtension($ext, $allowed) {
        if(in_array($ext, $allowed)) {
            return TRUE;
        } else {
            return FALSE;
        }
    }

    private function generateFilename($ext) {
        return uniqid("", true).'.'.$ext;
    }

    private function getExtension($name) {
        $ext = explode('.', $name);
        $ext = strtolower(end($ext));
        return $ext;
    }

    private function moveFile($tmpname, $ext) {
        $dir = "uploads/".$this->generateFilename($ext);
        if (move_uploaded_file($tmpname, $dir)) {
            return $dir;
        } else {
            return FALSE;
        }
    }

    private function filesToArray($files) {
        $new_files = array();
        foreach ($files['name'] as $key => $name) {
            $new_files[$key] = [
                'name' => $name,
                'tmp_name' => $files['tmp_name'][$key],
                'size' => $files['size'][$key],
                'error' => $files['error'][$key],
            ];
        }
        return $new_files;
    }

    private function handleFile($file) {
        $ext = $this->getExtension($file['name']);
        if($this->checkExtension($ext, ['msg','eml'])) {
            $move = $this->moveFile($file['tmp_name'], $ext);
            if($move !== FALSE) {
                return ['status'=>'success', 'path'=>$move];
            } else {
                return ['status'=>'fail', 'msg'=>'Failed to move file.'];
            }
        } else {
            return ['status'=>'fail', 'msg'=>'File has invalid extension.'];
        }
    }

    public function upload($file_array) {
        $processed = [
            'success' => array(),
            'failed' => array(),
        ];
        foreach($this->filesToArray($file_array) as $key => $file) {
            $handled = $this->handleFile($file);
            if($handled['status'] == 'success') {
                $processed['success'][] = ['path'=>$handled['path'], 'ext'=>$this->getExtension($file['name'])];
            } else {
                $processed['failed'][] = ['name'=>$file['name'], 'msg'=>$handled['msg']];
            }
        }
        return $processed;
    }
}

// $dir = readline('Specify Directory: '); // Request directory from User
// $dir = rtrim($dir, '/') . '/'; // Ensure trailing slash at end
// $files = array();
if(!empty($_FILES['files']['name'][0])) {

    $uploader = new fileHandler();

    $files = $_FILES['files'];
    $uploaded = $uploader->upload($files);

    if(count($uploaded['failed']) > 0) {
        echo 'The following files failed to upload:', "<br>";
        foreach ($uploaded['failed'] as $key => $val) {
            echo "[{$val['name']}] {$val['msg']}", "<br>";
        }
    }

    echo "====================== OUTPUTTING SENDER EMAIL ADDRESSES =======================", "<br>";
    echo "<br>";
    $extractor = new extractSender();
    foreach ($uploaded['success'] as $key => $val) {
        switch($val['ext']) {

            case 'eml':
                echo $extractor->parseEml($val['path']), "<br>";
            break;

            case 'msg':
                echo $extractor->parseMsg($val['path']), "<br>";
            break;
        }
        unlink($val['path']);
    }
    echo "<br>";
    echo "================================================================================";

} else {
    echo 'No files uploaded.';
}
// echo $dir, "<br>";
// $ole = $documentFactory->createFromFile('test.msg');
// $message = $messageFactory->parseMessage($ole);

// // raw properties are available from the "properties" member
// var_dump($message->properties['sender_email_address']);
// some properties have helper methods
// echo $message->getSender(), "<br>";
// echo $message->getBody(), "<br>";

// recipients and attachments are composed objects
// foreach ($message->getRecipients() as $recipient) {
//     // eg "To: John Smith <john.smith@example.com>
//     echo sprintf('%s: %s', $recipient->getType(), (string)$recipient), "<br>";
// }