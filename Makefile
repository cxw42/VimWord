all:
	bash -c './re2vba.pl --nodim vim-regex.txt |tee >(putclip)'
