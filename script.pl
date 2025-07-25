#!perl
use strict;
use warnings;

use utf8;
use Encode;

use Win32::GUI();
use PAR;
use File::Temp qw(tempfile);

use Excel::Writer::XLSX;

my $main = Win32::GUI::Window->new(
    -name       => 'Main',
    -size       => [800, 600],
    -minsize    => [800, 600],
    -text       => encode('gbk', '表格生成工具v1.0.0.0'),
    -background => 0xFFFFFF
);

my $tmp_icon;
if(my $file_tmp_icon = PAR::read_file('icon.ico')){
    my ($fh, $path) = tempfile(
        SUFFIX => '.ico',
        UNLINK => 1
    );
    binmode $fh;
    print $fh $file_tmp_icon;
    close $fh;
    $tmp_icon = $path;
}else{
    Win32::GUI::MessageBox(
        $main,
        encode('gbk', '图标文件损坏！'),
        encode('gbk', '错误'),
        0 + 16
    );
}

my $icon = Win32::GUI::Icon->new($tmp_icon) or Win32::GUI::MessageBox(
    $main,
    encode('gbk', '找不到图标文件！'),
    encode('gbk', '错误'),
    0 + 16
);
$main->SetIcon($icon);

my $main_width = $main->Width();
my $main_scale_width = $main->ScaleWidth();
my $main_height = $main->Height();
my $main_scale_height = $main->ScaleHeight();

my $desktop = Win32::GUI::GetDesktopWindow();
my $desktop_width = Win32::GUI::Width($desktop);
my $desktop_height = Win32::GUI::Height($desktop);

my $x = ($desktop_width - $main_width) / 2;
my $y = ($desktop_height - $main_height) / 2;
$main->Move($x, $y);

my $font_label_my = Win32::GUI::Font->new(
    -name => 'Arial',
    -size => 8
);

my $label_myname = $main->AddLabel(
    -text       => 'Created by Sleepyfish',
    -background => 0xFFFFFF,
    -font       => $font_label_my
);
my $label_myemail = $main->AddLabel(
    -text       => 'mackerel0203@outlook.com',
    -background => 0xFFFFFF,
    -font       => $font_label_my
);
my $label_myname_width = $label_myname->Width();
my $label_myname_height = $label_myname->Height();
my $label_myemail_width = $label_myemail->Width();
my $label_myemail_height = $label_myemail->Height();
$label_myname->Left($main_scale_width - $label_myname_width);
$label_myname->Top($main_scale_height - $label_myemail_height - $label_myname_height);
$label_myemail->Left($main_scale_width - $label_myemail_width);
$label_myemail->Top($main_scale_height - $label_myemail_height);

$main->AddLabel(
    -pos        => [30, 30],
    -text       => encode('gbk', '样品编号列表：'),
    -background => 0xFFFFFF
);

my $textfield_sample_num = $main->AddTextfield(
    -pos       => [124, 30],
    -size      => [100, 150],
    -multiline => 1,
    -vscroll   => 1,
    -hscroll   => 1
);

my $checkbox_fir_highlight = $main->AddCheckbox(
    -text       => encode('gbk', '首个样品高亮'),
    -pos        => [25, 60],
    -width      => 90,
    -name       => "Checkbox",
    -background => 0xFFFFFF,
    -checked    => 1
);

$main->AddLabel(
    -pos        => [280, 30],
    -text       => encode('gbk', '样品数量列表：'),
    -background => 0xFFFFFF
);

my $textfield_sample_rep = $main->AddTextfield(
    -pos       => [374, 30],
    -size      => [100, 150],
    -multiline => 1,
    -vscroll   => 1,
    -hscroll   => 1
);

my $checkbox_no_sample_rep = $main->AddCheckbox(
    -text       => encode('gbk', '不显示样品数'),
    -pos        => [275, 60],
    -width      => 90,
    -name       => "Checkbox",
    -background => 0xFFFFFF,
    -checked    => 0
);

$main->AddLabel(
    -pos        => [530, 30],
    -text       => encode('gbk', '末尾样品列表：'),
    -background => 0xFFFFFF
);

my $textfield_tail_sample = $main->AddTextfield(
    -pos       => [624, 30],
    -size      => [100, 150],
    -multiline => 1,
    -vscroll   => 1,
    -hscroll   => 1
);

my $checkbox_keep_blank = $main->AddCheckbox(
    -text       => encode('gbk', '保持末尾空白'),
    -pos        => [525, 60],
    -width      => 90,
    -name       => "Checkbox",
    -background => 0xFFFFFF,
    -checked    => 1
);

$main->AddLabel(
    -pos        => [522, 215],
    -text       => encode('gbk', '表格绘制起始点(行 列)：'),
    -background => 0xFFFFFF
);

my $textfield_init_row_pos = $main->AddTextfield(
    -pos  => [669, 210],
    -size => [25, 25],
    -text => 2
);

my $textfield_init_col_pos = $main->AddTextfield(
    -pos  => [699, 210],
    -size => [25, 25],
    -text => 1
);

$main->AddLabel(
    -pos        => [522, 270],
    -text       => encode('gbk', '        行范围：       '),
    -background => 0xFFFFFF
);

my $textfield_row_num_start = $main->AddTextfield(
    -pos  => [669, 265],
    -size => [25, 25],
    -text => 'A'
);

my $textfield_row_num_end = $main->AddTextfield(
    -pos  => [699, 265],
    -size => [25, 25],
    -text => 'H'
);

$main->AddLabel(
    -pos        => [522, 325],
    -text       => encode('gbk', '        列范围：       '),
    -background => 0xFFFFFF
);

my $textfield_col_num_start = $main->AddTextfield(
    -pos  => [669, 320],
    -size => [25, 25],
    -text => 1
);

my $textfield_col_num_end = $main->AddTextfield(
    -pos  => [699, 320],
    -size => [25, 25],
    -text => 12
);

$main->AddLabel(
    -pos        => [348, 215],
    -text       => encode('gbk', ' 字体大小： '),
    -background => 0xFFFFFF
);

my $textfield_font_size = $main->AddTextfield(
    -pos  => [429, 210],
    -size => [25, 25],
    -text => 9
);

$main->AddLabel(
    -pos        => [348, 270],
    -text       => encode('gbk', '末尾空白数：'),
    -background => 0xFFFFFF
);

my $textfield_tail_blank = $main->AddTextfield(
    -pos  => [429, 265],
    -size => [25, 25],
    -text => 0
);

my $checkbox_no_order_blank = $main->AddCheckbox(
    -text       => encode('gbk', '取消按照板序的空格'),
    -pos        => [340, 322],
    -width      => 132,
    -name       => "Checkbox",
    -background => 0xFFFFFF,
    -checked    => 0
);

$main->AddLabel(
    -pos        => [30, 215],
    -text       => encode('gbk', '  试验名称：  '),
    -background => 0xFFFFFF
);

my $textfield_customer_name = $main->AddTextfield(
    -pos  => [120, 210],
    -size => [160, 25]
);

$main->AddLabel(
    -pos        => [30, 270],
    -text       => encode('gbk', '表格起始编号：'),
    -background => 0xFFFFFF
);

my $textfield_table_num = $main->AddTextfield(
    -pos  => [120, 265],
    -size => [160, 25]
);

$main->AddLabel(
    -pos        => [30, 325],
    -text       => encode('gbk', ' 输出文件名： '),
    -background => 0xFFFFFF
);

my $textfield_output = $main->AddTextfield(
    -pos  => [120, 320],
    -size => [160, 25]
);

my $checkbox_horizontal_arr = $main->AddCheckbox(
    -text       => encode('gbk', '按照横向排列'),
    -pos        => [347, 367],
    -width      => 132,
    -name       => "Checkbox",
    -background => 0xFFFFFF,
    -checked    => 0
);

my $button_input = $main->AddButton(
    -name       => 'Input',
    -size       => [100, 30],
    -text       => encode('gbk', '导入'),
    -background => 0xFFFFFF
);

my $text_sample_num;
my $text_sample_rep;
my $customer_name;
my $table_num;
my $output;
my @sample_num;
my @sample_rep;
my $init_row_pos;
my $init_col_pos;
my $row_num_start;
my $row_num_end;
my @row_num;
my $col_num_start;
my $col_num_end;
my @col_num;
my $check_fir_highlight;
my $check_no_sample_rep;
my $font_size;
my $tail_blank;
my $text_tail_sample;
my @tail_sample;
my $check_keep_blank;
my $check_no_order_blank;
my $check_horizontal_arr;
sub Input_Click {
    my $tmp_text_sample_num = $textfield_sample_num->Text();
    my $tmp_text_sample_rep = $textfield_sample_rep->Text();
    my $tmp_customer_name = $textfield_customer_name->Text();
    my $tmp_table_num = $textfield_table_num->Text();
    my $tmp_output = $textfield_output->Text();
    my $tmp_init_row_pos = $textfield_init_row_pos->Text();
    my $tmp_init_col_pos = $textfield_init_col_pos->Text();
    my $tmp_row_num_start = $textfield_row_num_start->Text();
    my $tmp_row_num_end = $textfield_row_num_end->Text();
    my $tmp_col_num_start = $textfield_col_num_start->Text();
    my $tmp_col_num_end = $textfield_col_num_end->Text();
    my $tmp_check_fir_highlight = $checkbox_fir_highlight->Checked();
    my $tmp_check_no_sample_rep = $checkbox_no_sample_rep->Checked();
    my $tmp_font_size = $textfield_font_size->Text();
    my $tmp_tail_blank = $textfield_tail_blank->Text();
    my $tmp_text_tail_sample = $textfield_tail_sample->Text();
    my $tmp_check_keep_blank = $checkbox_keep_blank->Checked();
    my $tmp_check_no_order_blank = $checkbox_no_order_blank->Checked();
    my $tmp_check_horizontal_arr = $checkbox_horizontal_arr->Checked();
    if(!defined $tmp_text_sample_num || $tmp_text_sample_num !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '样品编号列表的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $tmp_text_sample_rep || $tmp_text_sample_rep !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '样品数量列表的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $tmp_customer_name || $tmp_customer_name !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '试验名称的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $tmp_table_num || $tmp_table_num !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '表格起始编号的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $tmp_output || $tmp_output !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '输出文件名的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif((!defined $tmp_init_row_pos || $tmp_init_row_pos !~ /\S/) || (!defined $tmp_init_col_pos || $tmp_init_col_pos !~ /\S/)){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '表格绘制起始点的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif((!defined $tmp_row_num_start || $tmp_row_num_start !~ /\S/) || (!defined $tmp_row_num_end || $tmp_row_num_end !~ /\S/)){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '行范围的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif((!defined $tmp_col_num_start || $tmp_col_num_start !~ /\S/) || (!defined $tmp_col_num_end || $tmp_col_num_end !~ /\S/)){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '列范围的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $tmp_font_size || $tmp_font_size !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '字体大小的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $tmp_tail_blank || $tmp_tail_blank !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '末尾空白数的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!$tmp_check_keep_blank && (!defined $tmp_text_tail_sample || $tmp_text_tail_sample !~ /\S/)){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '末尾样品列表的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_text_sample_num =~ /\t| /){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '样品编号列表的内容不能包含制表符或空格！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_text_sample_rep =~ /\t| /){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '样品数量列表的内容不能包含制表符或空格！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_text_tail_sample =~ /\t| /){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '末尾样品列表的内容不能包含制表符或空格！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_text_sample_rep =~ /[^\d\r\n]/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '样品数量列表的内容只能由大于0的纯数字组成！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_text_sample_rep =~ /^0/ || $tmp_text_sample_rep =~ /\r\n0/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '部分样品的数量为0或以0开头！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_output =~ /\<|\>|\:|\"|\/|\\|\||\*|\?/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '文件名中不能包含：< > : " / \ | * ?'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_table_num =~ /^([a-zA-Z]*0*)(.+)/ && $2 =~ /\D/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '表格编号只能由数字和字母组成，若有字母必须在编号开头！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_init_row_pos !~ /^\d+$/ || $tmp_init_col_pos !~ /^\d+$/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '表格绘制起始点只能由数字组成！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_col_num_start !~ /^\d+$/ || $tmp_col_num_end !~ /^\d+$/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '列范围只能由大于等于0的数字组成！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_col_num_start > $tmp_col_num_end){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '列范围起始值必须小于或等于终止值！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(($tmp_row_num_start !~ /^[A-Z]$/ || $tmp_row_num_end !~ /^[A-Z]$/) && ($tmp_row_num_start !~ /^[a-z]$/ || $tmp_row_num_end !~ /^[a-z]$/)){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '行范围只能由单个相同大小写的字母组成！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_row_num_start gt $tmp_row_num_end){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '行范围起始值必须小于或等于终止值！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_init_row_pos < 2){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '表格绘制起始点行坐标值必须大于等于2（至少为页码预留一行）！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_init_col_pos < 1){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '表格绘制起始点列坐标值必须大于等于1！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_font_size !~ /^\d+$/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '字体大小必须由大于等于1的数字组成！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_font_size < 1){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '字体大小必须由大于等于1的数字组成！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif($tmp_tail_blank !~ /^\d+$/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '末尾空白数必须由大于等于0的数字组成！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }else{
        $text_sample_num = $tmp_text_sample_num;
        $text_sample_rep = $tmp_text_sample_rep;
        $customer_name = decode('gbk', "$tmp_customer_name");
        $table_num = $tmp_table_num;
        $output = $tmp_output;
        $text_sample_num =~ s/\r\n/\t/g;
        @sample_num = split /\t/, $text_sample_num;
        $text_sample_rep =~ s/\r\n/\t/g;
        @sample_rep = split /\t/, $text_sample_rep;
        $init_row_pos = $tmp_init_row_pos - 1;
        $init_col_pos = $tmp_init_col_pos - 1;
        @col_num = ($tmp_col_num_start..$tmp_col_num_end);
        @row_num = ($tmp_row_num_start..$tmp_row_num_end);
        $check_fir_highlight = $tmp_check_fir_highlight;
        $check_no_sample_rep = $tmp_check_no_sample_rep;
        $font_size = $tmp_font_size;
        $tail_blank = $tmp_tail_blank;
        $check_keep_blank = $tmp_check_keep_blank;
        $check_no_order_blank = $tmp_check_no_order_blank;
        $check_horizontal_arr = $tmp_check_horizontal_arr;
        if(!$check_keep_blank){
            $text_tail_sample = $tmp_text_tail_sample;
            $text_tail_sample =~ s/\r\n/\t/g;
            my @temp_tail_sample = split /\t/, $text_tail_sample;
            if($#temp_tail_sample == $tail_blank - 1){
                @tail_sample = @temp_tail_sample;
            }else{
                Win32::GUI::MessageBox(
                    $main,
                    encode('gbk', '末尾样品数与末尾空白数不一致！'),
                    encode('gbk', '警告'),
                    0 + 48
                );
                return;
            }
        }
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '导入成功'),
            encode('gbk', '提示'),
            0 + 64
        );
        1;
    }
}

my $button_run = $main->AddButton(
    -name       => 'Run',
    -size       => [100, 30],
    -text       => encode('gbk', '运行'),
    -background => 0xFFFFFF
);

my $label_status = $main->AddLabel(
    -text       => '',
    -height     => 12,
    -foreground => 0x0000FF,
    -background => 0xFFFFFF
);
sub Run_Click {
    if(!@sample_num){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '样品编号列表的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!@sample_rep){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '样品数量列表的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $customer_name || $customer_name !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '试验名称的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $table_num || $table_num !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '表格起始编号的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $output || $output !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '输出文件名的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif((!defined $init_row_pos || $init_row_pos !~ /\S/) || (!defined $init_col_pos || $init_col_pos !~ /\S/)){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '表格绘制起始点的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!@row_num){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '行范围的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!@col_num){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '列范围的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $font_size || $font_size !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '字体大小的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!defined $tail_blank || $tail_blank !~ /\S/){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '末尾空白数的输入内容为空！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }elsif(!$check_keep_blank && !@tail_sample){
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '末尾样品列表的输入内容为空或数量有误！'),
            encode('gbk', '警告'),
            0 + 48
        );
        1;
    }else{
        Win32::GUI::MessageBox(
            $main,
            encode('gbk', '开始运行'),
            encode('gbk', '提示'),
            0 + 64
        );
        $label_status->Text(encode('gbk', '运行中'));
        $label_status->Width(36);
        my $label_status_width = $label_status->Width();
        $label_status->Left(($main_scale_width - $label_status_width) / 2);
        $label_status->Top($main_scale_height - 50);
        Win32::GUI::DoEvents();
################################################################################################################################################################################################
if($output =~ /(.+)\.xlsx$/){$output = $1;}

my $workbook = Excel::Writer::XLSX->new("${output}.xlsx");
unless(defined $workbook){
    Win32::GUI::MessageBox(
        $main,
        "${output}.xlsx" . encode('gbk', '已经打开或正在被占用，无法创建文件！'),
        encode('gbk', '错误'),
        0 + 16
    );
    $label_status->Text(encode('gbk', '运行错误'));
    $label_status->Width(48);
    $label_status_width = $label_status->Width();
    $label_status->Left(($main_scale_width - $label_status_width) / 2);
    $label_status->Top($main_scale_height - 50);
    return;
}

my $worksheet1_name = '样品排布表';
my $worksheet2_name = '竖_三列表';
my $worksheet3_name = '横_三列表';

my $worksheet1 = $workbook->add_worksheet($worksheet1_name);
my $worksheet2 = $workbook->add_worksheet($worksheet2_name);
my $worksheet3 = $workbook->add_worksheet($worksheet3_name);

unless($#sample_num == $#sample_rep){
    Win32::GUI::MessageBox(
        $main,
        encode('gbk', '样品编号与样品数量的个数不同！'),
        encode('gbk', '错误'),
        0 + 16
    );
    $label_status->Text(encode('gbk', '运行错误'));
    $label_status->Width(48);
    $label_status_width = $label_status->Width();
    $label_status->Left(($main_scale_width - $label_status_width) / 2);
    $label_status->Top($main_scale_height - 50);
    return;
}

my $sample_rep_len = 0;
for (@sample_rep){$sample_rep_len = length $_ if length $_ > $sample_rep_len;}

my $sample_count;
for (@sample_rep){$sample_count += $_;}
my $cell_count;
if(!$check_no_order_blank){
    $cell_count = ($#col_num + 1) * ($#row_num + 1) - 1 - $tail_blank;
}else{
    $cell_count = ($#col_num + 1) * ($#row_num + 1) - $tail_blank;
}

if($cell_count < 1){
    Win32::GUI::MessageBox(
        $main,
        encode('gbk', '末尾空白数太多，请减少空白的数量！'),
        encode('gbk', '错误'),
        0 + 16
    );
    $label_status->Text(encode('gbk', '运行错误'));
    $label_status->Width(48);
    $label_status_width = $label_status->Width();
    $label_status->Left(($main_scale_width - $label_status_width) / 2);
    $label_status->Top($main_scale_height - 50);
    return;
}

my $table_count;
if($sample_count % $cell_count != 0){
    $table_count = int($sample_count / $cell_count) + 1;
}else{
    $table_count = int($sample_count / $cell_count);
}

my $page_count;
if($table_count % 3 != 0){
    $page_count = int($table_count / 3) + 1;
}else{
    $page_count = int($table_count / 3);
}

my @table_num;
$table_num =~ /^([a-zA-Z]*0*)(.+)/;
for my $i ($2..$2 + $table_count - 1){
    my $j = "$1$i";
    push @table_num, $j;
}

my $sample_num_len = 0;
my $row_num_len = 0;
for (@sample_num){$sample_num_len = length $_ if length $_ > $sample_num_len;}
for (@row_num){$row_num_len = length $_ if length $_ > $row_num_len;}
my $column_len = $sample_num_len + $sample_rep_len - 0.5;
my $column_len_fir = $row_num_len + $sample_rep_len - 1;
for (0..0){
    my $col_start = $init_col_pos + 1;
    my $col_end = $init_col_pos + 1 + $#col_num;
    $worksheet1->set_column($col_start, $col_end, $column_len);
}
$worksheet1->set_column($init_col_pos, $init_col_pos, $column_len_fir);

my $format_table_num = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    bold         => 1,
    size         => 9,
    top          => 5,
    left         => 5,
    right        => 5,
    border_color => 'black'
);
my $format_row_num = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    bold         => 1,
    size         => 9,
    top          => 1,
    bottom       => 1,
    left         => 5,
    right        => 2,
    border_color => 'black'
);
my $format_row_num_last = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    bold         => 1,
    size         => 9,
    top          => 1,
    bottom       => 5,
    left         => 5,
    right        => 2,
    border_color => 'black'
);
my $format_col_num = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    bold         => 1,
    size         => 9,
    top          => 2,
    bottom       => 2,
    left         => 1,
    right        => 1,
    border_color => 'black'
);
my $format_col_num_last = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    bold         => 1,
    size         => 9,
    top          => 2,
    bottom       => 2,
    left         => 1,
    right        => 5,
    border_color => 'black'
);
my $format_sample_num = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    size         => $font_size,
    border       => 1,
    border_color => 'black'
);
my $format_sample_num_bottom = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    size         => $font_size,
    top          => 1,
    bottom       => 5,
    left         => 1,
    right        => 1,
    border_color => 'black'
);
my $format_sample_num_right = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    size         => $font_size,
    top          => 1,
    bottom       => 1,
    left         => 1,
    right        => 5,
    border_color => 'black'
);
my $format_sample_num_last = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    size         => $font_size,
    top          => 1,
    bottom       => 5,
    left         => 1,
    right        => 5,
    border_color => 'black'
);
my $format_sample_num_fir = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    size         => $font_size,
    bold         => 1,
    bg_color     => '#A5ECFF',
    border       => 1,
    border_color => 'black'
);
my $format_sample_num_fir_bottom = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    size         => $font_size,
    bold         => 1,
    bg_color     => '#A5ECFF',
    top          => 1,
    bottom       => 5,
    left         => 1,
    right        => 1,
    border_color => 'black' 
);
my $format_sample_num_fir_right = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    size         => $font_size,
    bold         => 1,
    bg_color     => '#A5ECFF',
    top          => 1,
    bottom       => 1,
    left         => 1,
    right        => 5,
    border_color => 'black' 
);
my $format_sample_num_fir_last = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    size         => $font_size,
    bold         => 1,
    bg_color     => '#A5ECFF',
    top          => 1,
    bottom       => 5,
    left         => 1,
    right        => 5,
    border_color => 'black'
);
my $format_page_num = $workbook->add_format(
    valign       => 'vcenter',
    align        => 'center',
    size         => 26,
    bold         => 1,
    font         => 'Microsoft YaHei'
);

my $format_left_border = $workbook->add_format(
    top          => 2,
    bottom       => 2,
    left         => 5,
    right        => 2,
    border_color => 'black'
);

for(my ($i, $row, $col) = (0, $init_row_pos + 1, $init_col_pos);$i <= $#table_num;$i++, $row = $row + $#row_num + 1 + 3){
    $worksheet1->write($row, $col, '', $format_left_border);
}

for(my ($i, $row_start, $col_start, $row_end, $col_end) = (0, $init_row_pos, $init_col_pos, $init_row_pos, $init_col_pos + $#col_num + 1);$i <= $#table_num;$i++, $row_start = $row_start + $#row_num + 1 + 3, $row_end = $row_end + $#row_num + 1 + 3){
    $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, $table_num[$i], $format_table_num);
}

for(my ($i, $row, $col) = (0, $init_row_pos + 1, $init_col_pos + 1);$i <= $#table_num;$i++, $row = $row + $#row_num + 1 + 3){
    for(my ($j, $k) = ($col, 0);$j <= $col + $#col_num;$j++, $k++){
        if($j == $col + $#col_num){
            $worksheet1->write($row, $j, $col_num[$k], $format_col_num_last);
        }else{
            $worksheet1->write($row, $j, $col_num[$k], $format_col_num);
        }
    }
}

for(my ($i, $row, $col) = (0, $init_row_pos + 2, $init_col_pos);$i <= $#table_num;$i++, $row = $row + $#row_num + 1 + 3){
    for(my ($j ,$k) = ($row, 0);$j <= $row + $#row_num;$j++, $k++){
        if($j == $row + $#row_num){
            $worksheet1->write($j, $col, $row_num[$k], $format_row_num_last);
        }else{
            $worksheet1->write($j, $col, $row_num[$k], $format_row_num);
        }
    }
}

my @sheet_title = ('板号', '孔号', '编号');
$worksheet2->write_row(0, 0, \@sheet_title);
$worksheet3->write_row(0, 0, \@sheet_title);

if(!$check_horizontal_arr){
    for(my ($i, $row, $col, $sample_num, $rep, $line_v, $line_h) = (0, $init_row_pos + 2, $init_col_pos + 1, 0, 1, 1, 1);$i <= $#table_num;$i++, $row = $row + $#row_num + 1 + 3, $line_h += ($#col_num + 1) * ($#row_num + 1)){
        my $order_blank_index;
        my $temp_order_blank_index = ($i + 1) % (($#col_num + 1) * ($#row_num + 1) - $tail_blank);
        if($temp_order_blank_index == 0){
            $order_blank_index = ($#col_num + 1) * ($#row_num + 1) - $tail_blank;
        }else{
            $order_blank_index = $temp_order_blank_index;
        }
        for(my ($j, $k, $count, $h, $count_limit, $tail_sample) = ($row, $col, 1, $line_h, ($#col_num + 1) * ($#row_num + 1) - $tail_blank, 0);$k <= $col + $#col_num;$j++, $rep++, $count++, $line_v++, $h += $#col_num + 1){
            my @array;
            if($sample_num > $#sample_num){
                if($count > $count_limit && !$check_keep_blank){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$tail_sample[$tail_sample]");
                    $worksheet2->write_row($line_v, 0, \@array);
                    $worksheet3->write_row($h, 0, \@array);
                    $tail_sample++;
                }else{
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, '', $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, '', $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, '', $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, '', $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", '');
                    $worksheet2->write_row($line_v, 0, \@array);
                    $worksheet3->write_row($h, 0, \@array);
                }
            }elsif(!$check_no_order_blank && $count == $order_blank_index){
                if($j == $row + $#row_num){
                    if($k == $col + $#col_num){
                        $worksheet1->write($j, $k, '', $format_sample_num_last);
                    }else{
                        $worksheet1->write($j, $k, '', $format_sample_num_bottom);
                    }
                }elsif($k == $col + $#col_num){
                    $worksheet1->write($j, $k, '', $format_sample_num_right);
                }else{
                    $worksheet1->write($j, $k, '', $format_sample_num);
                }
                @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", '');
                $worksheet2->write_row($line_v, 0, \@array);
                $worksheet3->write_row($h, 0, \@array);
                $rep -= 1;
            }elsif($count > $count_limit && $check_keep_blank){
                if($j == $row + $#row_num){
                    if($k == $col + $#col_num){
                        $worksheet1->write($j, $k, '', $format_sample_num_last);
                    }else{
                        $worksheet1->write($j, $k, '', $format_sample_num_bottom);
                    }
                }elsif($k == $col + $#col_num){
                    $worksheet1->write($j, $k, '', $format_sample_num_right);
                }else{
                    $worksheet1->write($j, $k, '', $format_sample_num);
                }
                @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", '');
                $worksheet2->write_row($line_v, 0, \@array);
                $worksheet3->write_row($h, 0, \@array);
                $rep -= 1;
            }elsif($count > $count_limit && !$check_keep_blank){
                if($j == $row + $#row_num){
                    if($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_last);
                    }else{
                        $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_bottom);
                    }
                }elsif($k == $col + $#col_num){
                    $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_right);
                }else{
                    $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num);
                }
                @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$tail_sample[$tail_sample]");
                $worksheet2->write_row($line_v, 0, \@array);
                $worksheet3->write_row($h, 0, \@array);
                $rep -= 1;
                $tail_sample++;
            }elsif($rep == 1){
                my $diff = $sample_rep_len - length $rep;
                my $full_rep = '0' x $diff . $rep;
                if($check_no_sample_rep && $check_fir_highlight){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_fir_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_fir_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_fir_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_fir);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]");
                }elsif(!$check_no_sample_rep && $check_fir_highlight){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_fir_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_fir_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_fir_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_fir);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]-$full_rep");
                }elsif($check_no_sample_rep && !$check_fir_highlight){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]");
                }elsif(!$check_no_sample_rep && !$check_fir_highlight){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]-$full_rep");
                }
                $worksheet2->write_row($line_v, 0, \@array);
                $worksheet3->write_row($h, 0, \@array);
            }else{
                my $diff = $sample_rep_len - length $rep;
                my $full_rep = '0' x $diff . $rep;
                if(!$check_no_sample_rep){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]-$full_rep");
                }elsif($check_no_sample_rep){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]");
                }
                $worksheet2->write_row($line_v, 0, \@array);
                $worksheet3->write_row($h, 0, \@array);
            }
            if($j == $row + $#row_num){$j = $row - 1;$h = $line_h - ($#col_num + 1) + $k - ($init_col_pos + 1) + 1;$k++;}
            unless($sample_num > $#sample_num){if($rep == $sample_rep[$sample_num]){$rep = 0;$sample_num++;}}
        }
    }
}elsif($check_horizontal_arr){
    for(my ($i, $row, $col, $sample_num, $rep, $line_v, $line_h) = (0, $init_row_pos + 2, $init_col_pos + 1, 0, 1, 1, 1);$i <= $#table_num;$i++, $row = $row + $#row_num + 1 + 3, $line_v += ($#col_num + 1) * ($#row_num + 1)){
        my $order_blank_index;
        my $temp_order_blank_index = ($i + 1) % (($#col_num + 1) * ($#row_num + 1) - $tail_blank);
        if($temp_order_blank_index == 0){
            $order_blank_index = ($#col_num + 1) * ($#row_num + 1) - $tail_blank;
        }else{
            $order_blank_index = $temp_order_blank_index;
        }
        for(my ($j, $k, $count, $v, $count_limit, $tail_sample) = ($row, $col, 1, $line_v, ($#col_num + 1) * ($#row_num + 1) - $tail_blank, 0);$j <= $row + $#row_num;$k++, $rep++, $count++, $line_h++, $v += $#row_num + 1){
            my @array;
            if($sample_num > $#sample_num){
                if($count > $count_limit && !$check_keep_blank){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$tail_sample[$tail_sample]");
                    $worksheet2->write_row($v, 0, \@array);
                    $worksheet3->write_row($line_h, 0, \@array);
                    $tail_sample++;
                }else{
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, '', $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, '', $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, '', $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, '', $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", '');
                    $worksheet2->write_row($v, 0, \@array);
                    $worksheet3->write_row($line_h, 0, \@array);
                }
            }elsif(!$check_no_order_blank && $count == $order_blank_index){
                if($j == $row + $#row_num){
                    if($k == $col + $#col_num){
                        $worksheet1->write($j, $k, '', $format_sample_num_last);
                    }else{
                        $worksheet1->write($j, $k, '', $format_sample_num_bottom);
                    }
                }elsif($k == $col + $#col_num){
                    $worksheet1->write($j, $k, '', $format_sample_num_right);
                }else{
                    $worksheet1->write($j, $k, '', $format_sample_num);
                }
                @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", '');
                $worksheet2->write_row($v, 0, \@array);
                $worksheet3->write_row($line_h, 0, \@array);
                $rep -= 1;
            }elsif($count > $count_limit && $check_keep_blank){
                if($j == $row + $#row_num){
                    if($k == $col + $#col_num){
                        $worksheet1->write($j, $k, '', $format_sample_num_last);
                    }else{
                        $worksheet1->write($j, $k, '', $format_sample_num_bottom);
                    }
                }elsif($k == $col + $#col_num){
                    $worksheet1->write($j, $k, '', $format_sample_num_right);
                }else{
                    $worksheet1->write($j, $k, '', $format_sample_num);
                }
                @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", '');
                $worksheet2->write_row($v, 0, \@array);
                $worksheet3->write_row($line_h, 0, \@array);
                $rep -= 1;
            }elsif($count > $count_limit && !$check_keep_blank){
                if($j == $row + $#row_num){
                    if($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_last);
                    }else{
                        $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_bottom);
                    }
                }elsif($k == $col + $#col_num){
                    $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num_right);
                }else{
                    $worksheet1->write($j, $k, "$tail_sample[$tail_sample]", $format_sample_num);
                }
                @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$tail_sample[$tail_sample]");
                $worksheet2->write_row($v, 0, \@array);
                $worksheet3->write_row($line_h, 0, \@array);
                $rep -= 1;
                $tail_sample++;
            }elsif($rep == 1){
                my $diff = $sample_rep_len - length $rep;
                my $full_rep = '0' x $diff . $rep;
                if($check_no_sample_rep && $check_fir_highlight){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_fir_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_fir_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_fir_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_fir);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]");
                }elsif(!$check_no_sample_rep && $check_fir_highlight){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_fir_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_fir_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_fir_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_fir);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]-$full_rep");
                }elsif($check_no_sample_rep && !$check_fir_highlight){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]");
                }elsif(!$check_no_sample_rep && !$check_fir_highlight){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]-$full_rep");
                }
                $worksheet2->write_row($v, 0, \@array);
                $worksheet3->write_row($line_h, 0, \@array);
            }else{
                my $diff = $sample_rep_len - length $rep;
                my $full_rep = '0' x $diff . $rep;
                if(!$check_no_sample_rep){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]-$full_rep", $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]-$full_rep");
                }elsif($check_no_sample_rep){
                    if($j == $row + $#row_num){
                        if($k == $col + $#col_num){
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_last);
                        }else{
                            $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_bottom);
                        }
                    }elsif($k == $col + $#col_num){
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num_right);
                    }else{
                        $worksheet1->write($j, $k, "$sample_num[$sample_num]", $format_sample_num);
                    }
                    @array = ($table_num[$i], "$row_num[$j - (($#row_num + 1 + 3) * $i) - 2 - $init_row_pos]"."$col_num[$k - 1 - $init_col_pos]", "$sample_num[$sample_num]");
                }
                $worksheet2->write_row($v, 0, \@array);
                $worksheet3->write_row($line_h, 0, \@array);
            }
            if($k == $col + $#col_num){$k = $col - 1;$v = $line_v - ($#row_num + 1) + $j - ($init_row_pos + 2) + 1 - ($#row_num + 1 + 3) * $i;$j++;}
            unless($sample_num > $#sample_num){if($rep == $sample_rep[$sample_num]){$rep = 0;$sample_num++;}}
        }
    }
}

if($#col_num >= 4){
    my $s;
    if(($#col_num + 1 - 2) % 3 == 0){
        $s = ($#col_num + 1 - 2) / 3;
        for(my ($i, $row_start, $col_start, $row_end, $col_end) = (1, $init_row_pos - 1, $init_col_pos + 1, $init_row_pos - 1, $init_col_pos + $s * 2);$i <= $page_count;$i++, $row_start = $row_start + ($#row_num + 1 + 3) * 3, $row_end = $row_end + ($#row_num + 1 + 3) * 3){
            $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, $customer_name, $format_page_num);
        }
        for(my ($i, $row_start, $col_start, $row_end, $col_end) = (1, $init_row_pos - 1, $init_col_pos + $s * 2 + 1 + 1, $init_row_pos - 1, $init_col_pos + $s * 2 + 1 + $s);$i <= $page_count;$i++, $row_start = $row_start + ($#row_num + 1 + 3) * 3, $row_end = $row_end + ($#row_num + 1 + 3) * 3){
            if($col_end == $col_start){
                $worksheet1->write($row_start, $col_end, "第 $i 页", $format_page_num);
            }else{
                $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, "第 $i 页", $format_page_num);
            }
        }
    }elsif(($#col_num + 1 - 2) % 3 == 1){
        $s = int(($#col_num + 1 - 2) / 3);
        for(my ($i, $row_start, $col_start, $row_end, $col_end) = (1, $init_row_pos - 1, $init_col_pos + 1, $init_row_pos - 1, $init_col_pos + $s * 2 + 1);$i <= $page_count;$i++, $row_start = $row_start + ($#row_num + 1 + 3) * 3, $row_end = $row_end + ($#row_num + 1 + 3) * 3){
            $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, $customer_name, $format_page_num);
        }
        for(my ($i, $row_start, $col_start, $row_end, $col_end) = (1, $init_row_pos - 1, $init_col_pos + $s * 2 + 1 + 1 + 1, $init_row_pos - 1, $init_col_pos + $s * 2 + 1 + 1 + $s);$i <= $page_count;$i++, $row_start = $row_start + ($#row_num + 1 + 3) * 3, $row_end = $row_end + ($#row_num + 1 + 3) * 3){
            if($col_end == $col_start){
                $worksheet1->write($row_start, $col_end, "第 $i 页", $format_page_num);
            }else{
                $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, "第 $i 页", $format_page_num);
            }
        }
    }elsif(($#col_num + 1 - 2) % 3 == 2){
        $s = int(($#col_num + 1 - 2) / 3);
        for(my ($i, $row_start, $col_start, $row_end, $col_end) = (1, $init_row_pos - 1, $init_col_pos + 1, $init_row_pos - 1, $init_col_pos + $s * 2 + 1);$i <= $page_count;$i++, $row_start = $row_start + ($#row_num + 1 + 3) * 3, $row_end = $row_end + ($#row_num + 1 + 3) * 3){
            $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, $customer_name, $format_page_num);
        }
        for(my ($i, $row_start, $col_start, $row_end, $col_end) = (1, $init_row_pos - 1, $init_col_pos + $s * 2 + 1 + 1 + 1, $init_row_pos - 1, $init_col_pos + $s * 2 + 1 + 1 + $s + 1);$i <= $page_count;$i++, $row_start = $row_start + ($#row_num + 1 + 3) * 3, $row_end = $row_end + ($#row_num + 1 + 3) * 3){
            $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, "第 $i 页", $format_page_num);
        }
    }
}elsif($#col_num == 3){
    for(my ($i, $row_start, $col_start, $row_end, $col_end) = (1, $init_row_pos - 1, $init_col_pos, $init_row_pos - 1, $init_col_pos + 1 + 1);$i <= $page_count;$i++, $row_start = $row_start + ($#row_num + 1 + 3) * 3, $row_end = $row_end + ($#row_num + 1 + 3) * 3){
        $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, $customer_name, $format_page_num);
    }
    for(my ($i, $row_start, $col_start, $row_end, $col_end) = (1, $init_row_pos - 1, $init_col_pos + 1 + 1 + 1, $init_row_pos - 1, $init_col_pos + 1 + 1 + 1 + 1);$i <= $page_count;$i++, $row_start = $row_start + ($#row_num + 1 + 3) * 3, $row_end = $row_end + ($#row_num + 1 + 3) * 3){
        $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, "第 $i 页", $format_page_num);
    }
}elsif($#col_num == 2){
    for(my ($i, $row_start, $col_start, $row_end, $col_end) = (1, $init_row_pos - 1, $init_col_pos, $init_row_pos - 1, $init_col_pos + 1 + 1);$i <= $page_count;$i++, $row_start = $row_start + ($#row_num + 1 + 3) * 3, $row_end = $row_end + ($#row_num + 1 + 3) * 3){
        $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, $customer_name, $format_page_num);
    }
    for(my ($i, $row, $col) = (1, $init_row_pos - 1, $init_col_pos + 1 + 1 + 1);$i <= $page_count;$i++, $row = $row + ($#row_num + 1 + 3) * 3){
        $worksheet1->write($row, $col, "第 $i 页", $format_page_num);
    }
}elsif($#col_num == 1){
    for(my ($i, $row_start, $col_start, $row_end, $col_end) = (1, $init_row_pos - 1, $init_col_pos, $init_row_pos - 1, $init_col_pos + 1);$i <= $page_count;$i++, $row_start = $row_start + ($#row_num + 1 + 3) * 3, $row_end = $row_end + ($#row_num + 1 + 3) * 3){
        $worksheet1->merge_range($row_start, $col_start, $row_end, $col_end, $customer_name, $format_page_num);
    }
    for(my ($i, $row, $col) = (1, $init_row_pos - 1, $init_col_pos + 1 + 1);$i <= $page_count;$i++, $row = $row + ($#row_num + 1 + 3) * 3){
        $worksheet1->write($row, $col, "第 $i 页", $format_page_num);
    }
}elsif($#col_num == 0){
    for(my ($i, $row, $col) = (1, $init_row_pos - 1, $init_col_pos);$i <= $page_count;$i++, $row = $row + ($#row_num + 1 + 3) * 3){
        $worksheet1->write($row, $col, $customer_name, $format_page_num);
    }
    for(my ($i, $row, $col) = (1, $init_row_pos - 1, $init_col_pos + 1);$i <= $page_count;$i++, $row = $row + ($#row_num + 1 + 3) * 3){
        $worksheet1->write($row, $col, "第 $i 页", $format_page_num);
    }
}
################################################################################################################################################################################################
        if($workbook->close()){
            $label_status->Text(encode('gbk', '运行结束'));
        }else{
            Win32::GUI::MessageBox(
                $main,
                encode('gbk', '无法正确关闭文件，运行错误！'),
                encode('gbk', '错误'),
                0 + 16
            );
            $label_status->Text(encode('gbk', '运行错误'));
        }
        $label_status->Width(48);
        $label_status_width = $label_status->Width();
        $label_status->Left(($main_scale_width - $label_status_width) / 2);
        $label_status->Top($main_scale_height - 50);
        1;
    }
}

$main->Show();
Win32::GUI::Dialog();

sub Main_Terminate {
    -1;
}

sub Main_Resize {
    my $main_scale_width = $main->ScaleWidth();
    my $main_scale_height = $main->ScaleHeight();

    my $button_run_width = $button_run->Width();
    $button_run->Left(($main_scale_width - $button_run_width) / 2);
    $button_run->Top($main_scale_height - 100);

    my $button_input_width = $button_input->Width();
    $button_input->Left(($main_scale_width - $button_input_width) / 2);
    $button_input->Top($main_scale_height - 150);

    if(defined $label_status){
        my $label_status_width = $label_status->Width();
        $label_status->Left(($main_scale_width - $label_status_width) / 2);
        $label_status->Top($main_scale_height - 50);
    }

    my $label_myname_width = $label_myname->Width();
    my $label_myname_height = $label_myname->Height();
    my $label_myemail_width = $label_myemail->Width();
    my $label_myemail_height = $label_myemail->Height();
    $label_myname->Left($main_scale_width - $label_myname_width);
    $label_myname->Top($main_scale_height - $label_myemail_height - $label_myname_height);
    $label_myemail->Left($main_scale_width - $label_myemail_width);
    $label_myemail->Top($main_scale_height - $label_myemail_height);
}
