BEGIN {
	FS = " " 
}

{
	if ( $1 ~ /Address/ )
	{
		IPaddress=$NF
	}
}


END { 
	printf("%s",IPaddress)
}