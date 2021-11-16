#include<unistd.h>
#include<stdio.h>
#include<stdlib.h>

int main()
{

pid_t pid;
pid = fork();
if(pid == 0)
{
	printf(" child : %d", pid);
	return 0;
}
else
{
	printf("father %d", pid);
}
return 0;
}
