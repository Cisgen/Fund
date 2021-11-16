#include <iostream>
#include <curl/curl.h>
#include <stdlib.h>
#include <string.h>
#include <string>
#include <map>
#include <vector>
#include <errno.h>

#define URL_NAME 0
#define URL_OPEN_PRICE 1
#define URL_CLOSE_PRICE 2
#define URL_CUR_PRICE 3
#define URL_HIGH_PRICE 4
#define URL_LOW_PRICE 5
#define URL_MAX 6

#define URL_CONFIG_NUMBER 0
#define URL_CONFIG_NAME 1
#define URL_CONFIG_BUY 2
#define URL_CONFIG_VALUE 3
#define URL_CONFIG_TRIGGER 4
#define URL_CONFIG_TRIGGERVAL 5
#define MAX_FUND_ID 7

#define PATH_CONFIG "./fund.config"
using namespace std;
struct FundVal
{
	string sName;
	float fCurVal;
	float fOpenVal;
	float fYesVal;
	float fHighVal;
	float fLowVal;
};

struct FundValConfig
{
	//购买份额
	int iBuyCnt;
	//成本价
	float fValue;
	//是否触发
	bool bTrigger;
	//触发值
	float fTriggerVal;
};

void http_callback(char* ptr, size_t size, size_t nmemb, map<string, FundValConfig>* userdata)
{
	string res(ptr);
	//对URL结果按照每个代码进行处理分割
	map<string, FundVal> map_res;
	const char* tok = "\",";
	map<int, string> allUrlRet;
	int ilen = res.length();
	int iBegin = 0;
	for (int i = 0; i < ilen; i++)
	{
		char cChar = res.at(i);
		if (cChar == '=')
		{
			//六位发财神秘代码
			iBegin = i;
		}
		else if (cChar == ';')
		{
			string stKeyName(res.begin() + iBegin - MAX_FUND_ID + 1, res.begin() + iBegin);
			string newstring(res.begin() + iBegin + 1, res.begin() + i);
			char* p = const_cast<char*>(newstring.c_str());

			//对每一行结果根据分隔符分割
			char* pToken = strtok(p, tok);
			FundVal stFund;
			int iRround = 0;
			while(pToken && iRround < URL_MAX)
			{
				switch (iRround)
				{
				case URL_NAME:
					stFund.sName.assign(pToken);
					break;
				case URL_CUR_PRICE:
					stFund.fCurVal = atof(pToken);
					break;
				case URL_CLOSE_PRICE:
					stFund.fYesVal = atof(pToken);
					break;
				case URL_HIGH_PRICE:
					stFund.fHighVal = atof(pToken);
					break;
				case URL_LOW_PRICE:
					stFund.fLowVal = atof(pToken);
					break;
				case URL_OPEN_PRICE:
					stFund.fOpenVal = atof(pToken);
					break;
				default:
					break;
				}

				pToken = strtok(NULL, tok);
				iRround++;
			}

			map_res.insert(make_pair(stKeyName, stFund));
		}
	}

	if (map_res.size() != userdata->size())
	{
		return;
	}

	printf("%10s\t%4s\t%6s\t%7s\t%7s\t%7s\n", "名称", "价格", "涨幅", "涨幅比", "盈利比", "盈利金额");
	float fTatal = 0;
	float fToday = 0;
	for (map<string, FundVal>::iterator it = map_res.begin(); it != map_res.end(); it++)
	{
		float fAddRate = 0;
		float fAddVal = 0;
		float fDiff = it->second.fCurVal - it->second.fYesVal;
		map<string, FundValConfig>::iterator itCfg = (*userdata).find(it->first);
		if (itCfg != (*userdata).end() && itCfg->second.fValue > 0)
		{
			fAddRate = (it->second.fCurVal - itCfg->second.fValue) * 100 / itCfg->second.fValue;
			fAddVal = (it->second.fCurVal - itCfg->second.fValue) * itCfg->second.iBuyCnt * 100;
			fTatal += fAddVal;
			fToday += itCfg->second.iBuyCnt * 100 * fDiff;
		}

		printf("%10s\t%6.2f\t", it->second.sName.c_str(), it->second.fCurVal);
		
		if (fDiff > 0)
		{
			printf("\033[31m%+7.2f\t%+7.2f\033[0m\t", fDiff , fDiff * 100 / it->second.fYesVal);
		}
		else if (fDiff < 0)
		{
			printf("\033[32m%+7.2f\t%+7.2f\033[0m\t", fDiff , fDiff * 100 / it->second.fYesVal);
		}
		else
		{
			printf("%+7.2f\t%+7.2f\t", fAddRate, fAddVal);
		}

		if (fAddVal > 0)
		{
			printf("\033[31m%+7.2f\t%8.2f\033[0m\n", fAddRate, fAddVal);
		}
		else if (fAddVal < 0)
		{
			printf("\033[32m%+7.2f\t%8.2f\033[0m\n", fAddRate, fAddVal);
		}
		else
		{
			printf("%+7.2f\t%8.2f\n", fAddRate, fAddVal);
		}
	}

	if (fTatal != 0 || fToday != 0)
	{
		if (fTatal > 0)
		{
			printf("当前总盈利金额:\033[31m %+7.2f\033[0m\t", fTatal);
		}
		else
		{
			printf("当前总盈利金额:\033[32m %+7.2f\033[0m\t", fTatal);
		}

		if (fToday > 0)
		{
			printf("本日盈利金额:\033[31m %+7.2f\033[0m\n", fToday);
		}
		else
		{
			printf("本日盈利金额:\033[32m %+7.2f\033[0m\n", fToday);
		}
	}
}

int main()
{
	//读取配置
	FILE *pFile = fopen(PATH_CONFIG, "rb");
	if (!pFile)
	{
		printf("Failed to open file:%s, error: %s", PATH_CONFIG, strerror(errno));
		return -1;
	}
	
	fseek(pFile, 0, SEEK_SET);
	rewind (pFile);

	char* pLine = NULL;
	size_t iLen = 0;
	string Url = "http://hq.sinajs.cn/list=";
	map<string, FundValConfig> mapConfig;
	bool bFirst = true;
	while (!feof(pFile) && !ferror(pFile))
	{
		getline(&pLine, &iLen, pFile);
		string stName;
		FundValConfig stFundConfig;
		if (pLine)
		{
			//数字开头
			if (pLine[0] < '0' || pLine[0] > '9')
			{
				continue;
			}

			const char* tok = ",";
			char* pToken = strtok(pLine, tok);
			int iCount = 0;
			while(pToken)
			{
				switch (iCount)
				{
				case URL_CONFIG_NUMBER:
					stName.assign(pToken);
					if (!bFirst)
					{
						Url.append(tok);
					}

					if (pToken[0] == '3' || pToken[0] == '0')
					{
						Url.append("sz");
					}
					else if (pToken[0] == '6')
					{
						Url.append("sh");
					}
					else
					{
						continue;
					}

					Url.append(pToken);
					bFirst = false;
					break;
				case URL_CONFIG_BUY:
					stFundConfig.iBuyCnt = atof(pToken);
					break;
				case URL_CONFIG_VALUE:
					stFundConfig.fValue = atof(pToken);
					break;
				case URL_CONFIG_TRIGGER:
					stFundConfig.bTrigger = atof(pToken);
					break;
				case URL_CONFIG_TRIGGERVAL:
					stFundConfig.fTriggerVal = atof(pToken);
					break;
				default:
					break;
				}

				pToken = strtok(NULL, tok);
				iCount++;
			}

			mapConfig.insert(make_pair(stName, stFundConfig));
		}
	}

	fclose(pFile);
	CURL* curl = curl_easy_init();
	if (curl)
	{
		while (1)
		{
			curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, &http_callback);
			curl_easy_setopt(curl, CURLOPT_WRITEDATA, &mapConfig);
			curl_easy_setopt(curl, CURLOPT_URL, Url.c_str());
			curl_easy_setopt(curl, CURLOPT_TIMEOUT, 120);

			//调用程序
			curl_easy_perform(curl);
			//daemon(1, 1);
			sleep(5);
		}
		
		//清除curl操作
		curl_easy_cleanup(curl);
	}

	return 0;
}