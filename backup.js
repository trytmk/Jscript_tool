//
//  �����N���b�N�Ŏw�肵���t�@�C���̃o�b�N�A�b�v�����js
//  ���Â��āubackup.js�v
//
//--- �ϐ����A�ݒ肷�邱��---------------------------------
//
// �o�b�N�A�b�N��t�H���_�ijs�̂���t�H���_�̒����j��ݒ�
//
var backupD = "backup";
//
// �o�b�N�A�b�N����t�@�C����ݒ�
//
var targetFile = "backup.js";
//
//------------------------------------------------------------
//  Shell�֘A�̑����񋟂���I�u�W�F�N�g���擾
var sh = new ActiveXObject( "WScript.Shell" );
//  �t�@�C���֘A�̑����񋟂���I�u�W�F�N�g���擾
var fs = new ActiveXObject( "Scripting.FileSystemObject" );
// �J�����g�t�H���_�ijs�̂���t�H���_�j���擾
var CurrentD = sh.CurrentDirectory +"\\";
// �o�b�N�A�b�N��t�H���_�̃t���p�X�𐶐�
var targetFP = CurrentD+backupD+"\\";

//------------------------------------------------------------
//�@�o�b�N�A�b�v��t�H���_�̏���
//  �ˎw�肵���t�H���_�����݂��Ă邩���m�F����B
// �@ https://jscript.zouri.jp/Source/FileFolderCtrl.html
if( fs.FolderExists(targetFP) ){
    //����Ƃ��͉������Ȃ�
}else{
    //�Ȃ���΃t�H���_�[�����
    fs.CreateFolder( targetFP )
    WScript.Echo( "�o�b�N�A�b�v�t�H���_���쐬���܂����B");
};
//------------------------------------------------------------
// YYMMDDHHmm_�^�œ������擾
// ��https://goat-inc.co.jp/blog/2584/
var dt  = new Date();
 var ymdhm =         ("00" + (dt.getFullYear())).slice(-2);
 var ymdhm = ymdhm + ("00" + (dt.getMonth()+1)).slice(-2);
 var ymdhm = ymdhm + ("00" + (dt.getDate())).slice(-2);
 var ymdhm = ymdhm + ("00" + (dt.getHours())).slice(-2);
 var ymdhm = ymdhm + ("00" + (dt.getMinutes())).slice(-2);
 var ymdhm = ymdhm + "_";

//
//�@�t�@�C���R�s�[
fs.CopyFile( targetFile, targetFP+ymdhm+targetFile);

//�@
WScript.Echo( "�o�b�N�A�b�v���܂����B�� .\\" + backupD + "\\" + ymdhm + targetFile );

//  �I�u�W�F�N�g�����
fs = null;
sh = null;
