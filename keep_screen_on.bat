@echo off

REM �Ǘ��Ҍ����Ŏ��s����Ă��邩�m�F
whoami /groups | find "S-1-16-12288" > nul
if errorlevel 1 (
    echo ���̃X�N���v�g�͊Ǘ��Ҍ����Ŏ��s����K�v������܂��B
    pause
    exit /b
)

REM ���݂̃X���[�v�ݒ�����O�ɏo��
powercfg /query > sleep_log.txt
if not exist sleep_log.txt (
    echo powercfg �R�}���h�̎��s�Ɏ��s���܂����B
    pause
    exit /b
)

REM ���O����AC�d����DC�d���̃X���[�v�ݒ�l���擾
for /f "tokens=5 delims=: " %%a in ('findstr /i /c:"���݂� AC �d���ݒ�̃C���f�b�N�X" sleep_log.txt') do set SLEEP_AC_ORIGINAL=%%a
for /f "tokens=5 delims=: " %%a in ('findstr /i /c:"���݂� DC �d���ݒ�̃C���f�b�N�X" sleep_log.txt') do set SLEEP_DC_ORIGINAL=%%a

REM �l���擾�ł������m�F
if "%SLEEP_AC_ORIGINAL%"=="" (
    echo AC�d���̃X���[�v�ݒ�l���擾�ł��܂���ł����B
    type sleep_log.txt
    pause
    exit /b
)
if "%SLEEP_DC_ORIGINAL%"=="" (
    echo DC�d���̃X���[�v�ݒ�l���擾�ł��܂���ł����B
    type sleep_log.txt
    pause
    exit /b
)

REM �X���[�v�𖳌���
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0

echo �X���[�v�𖳌������܂����B�I������ɂ�Ctrl+C�������Ă��������B

:LOOP
REM �������p��
ping -n 2 127.0.0.1 > nul
GOTO LOOP

:EXIT
REM ���̐ݒ�𕜌�
powercfg /change standby-timeout-ac %SLEEP_AC_ORIGINAL%
powercfg /change standby-timeout-dc %SLEEP_DC_ORIGINAL%

echo �X���[�v�̐ݒ�����ɖ߂��܂����B
pause
