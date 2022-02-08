using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooilAnalysis
{
    internal class ColumnDictionary
    {
        static Dictionary<string, string> dict = new Dictionary<string, string>();

        public ColumnDictionary()
        {
            dict.Add("work_date", "일자");
            dict.Add("sqen_numb", "순번");
            dict.Add("proc_date", "생산일자");
            dict.Add("proc_time", "생산시간");
            dict.Add("ds_data001", "메인속도표시");
            dict.Add("ds_data002", "텐터폭");
            dict.Add("ds_data003", "입구폭");
            dict.Add("ds_data004", "중간폭");
            dict.Add("ds_data005", "챔바1설정온도1");
            dict.Add("ds_data006", "챔바1설정온도2");
            dict.Add("ds_data007", "챔바2설정온도1");
            dict.Add("ds_data008", "챔바2설정온도2");
            dict.Add("ds_data009", "챔바3설정온도1");
            dict.Add("ds_data010", "챔바3설정온도2");
            dict.Add("ds_data011", "챔바4설정온도1");
            dict.Add("ds_data012", "챔바4설정온도2");
            dict.Add("ds_data013", "챔바5설정온도1");
            dict.Add("ds_data014", "챔바5설정온도2");
            dict.Add("ds_data015", "챔바6설정온도1");
            dict.Add("ds_data016", "챔바6설정온도2");
            dict.Add("ds_data017", "챔바7설정온도1");
            dict.Add("ds_data018", "챔바7설정온도2");
            dict.Add("ds_data019", "챔바1현재온도1");
            dict.Add("ds_data020", "챔바1현재온도2");
            dict.Add("ds_data021", "챔바2현재온도1");
            dict.Add("ds_data022", "챔바2현재온도2");
            dict.Add("ds_data023", "챔바3현재온도1");
            dict.Add("ds_data024", "챔바3현재온도2");
            dict.Add("ds_data025", "챔바4현재온도1");
            dict.Add("ds_data026", "챔바4현재온도2");
            dict.Add("ds_data027", "챔바5현재온도1");
            dict.Add("ds_data028", "챔바5현재온도2");
            dict.Add("ds_data029", "챔바6현재온도1");
            dict.Add("ds_data030", "챔바6현재온도2");
            dict.Add("ds_data031", "챔바7현재온도1");
            dict.Add("ds_data032", "챔바7현재온도2");
            dict.Add("ds_data033", "포온도3현재온도");
            dict.Add("ds_data034", "포온도4현재온도");
            dict.Add("ds_data035", "포온도5현재온도");
            dict.Add("ds_data036", "포온도6현재온도");
            dict.Add("ds_data037", "포온도7현재온도");
            dict.Add("ds_data038", "챔바1rpm 팬설정속도");
            dict.Add("ds_data039", "챔바2rpm 팬설정속도");
            dict.Add("ds_data040", "챔바3rpm 팬설정속도");
            dict.Add("ds_data041", "챔바4rpm 팬설정속도");
            dict.Add("ds_data042", "챔바5rpm 팬설정속도");
            dict.Add("ds_data043", "챔바6rpm 팬설정속도");
            dict.Add("ds_data044", "챔바7rpm 팬설정속도");
            dict.Add("ds_data045", "챔바1rpm 팬현재속도");
            dict.Add("ds_data046", "챔바2rpm 팬현재속도");
            dict.Add("ds_data047", "챔바3rpm 팬현재속도");
            dict.Add("ds_data048", "챔바4rpm 팬현재속도");
            dict.Add("ds_data049", "챔바5rpm 팬현재속도");
            dict.Add("ds_data050", "챔바6rpm 팬현재속도");
            dict.Add("ds_data051", "챔바7rpm 팬현재속도");
            dict.Add("ds_data052", "오버피드값");
            dict.Add("ds_data053", "피드인값");
            dict.Add("ds_data054", "배기1온도");
            dict.Add("ds_data055", "배기습도 설정값");
            dict.Add("ds_data056", "배기습도 현재값");
            dict.Add("ds_data057", "배기팬 설정속도1");
            dict.Add("ds_data058", "배기팬 현재속도1");
            dict.Add("ds_data059", "배기팬 설정속도1");
            dict.Add("ds_data060", "배기팬 현재속도1");
            dict.Add("ds_data061", "배기RPM");
            dict.Add("ds_data062", "덴타 출구폭");
            dict.Add("ds_data063", "드웰시간");
            dict.Add("ds_data064", "노출시간");
            dict.Add("ds_data065", "드웰온도");
            dict.Add("ds_data066", "지령속도");
            dict.Add("ds_data067", "오버피드롤 설정값");
            dict.Add("ds_data068", "피드인롤 설정값");
            dict.Add("ds_data069", "핀닝롤 설정값");
            dict.Add("ds_data070", "출구 딜리버리값");
            dict.Add("ds_data071", "출구 냉각실린더값");
            dict.Add("ds_data072", "출구 밧칭값");
            dict.Add("ds_data073", "출구 딜리버리 설정값");
            dict.Add("ds_data074", "출구 냉각실린더 설정값");
            dict.Add("ds_data075", "출구 밧칭 설정값");
            dict.Add("ds_data076", "텐다 운전ON");
            dict.Add("ds_data077", "텐다 준비ON");
            dict.Add("ds_data078", "망글 운전ON");
            dict.Add("ds_data079", "정련기 평균 상전압");
            dict.Add("ds_data080", "정련기 평균 선간전압");
            dict.Add("ds_data081", "정련기 평균 상전류");
            dict.Add("ds_data082", "정련기 총합유효전력");
            dict.Add("ds_data083", "정련기 총합무효전력");
            dict.Add("ds_data084", "정련기 역률");
            dict.Add("ds_data085", "정련기 당월누적전력량");
            dict.Add("ds_data086", "직거 평균 상전압");
            dict.Add("ds_data087", "직거 평균 선간전압");
            dict.Add("ds_data088", "직거 평균 상전류");
            dict.Add("ds_data089", "직거 총합유효전력");
            dict.Add("ds_data090", "직거 총합무효전력");
            dict.Add("ds_data091", "직거 역률");
            dict.Add("ds_data092", "직거 당월누적전력량");
            dict.Add("ds_data093", "래피드 평균 상전압");
            dict.Add("ds_data094", "래피드 평균 선간전압");
            dict.Add("ds_data095", "래피드 평균 상전류");
            dict.Add("ds_data096", "래피드 총합유효전력");
            dict.Add("ds_data097", "래피드 총합무효전력");
            dict.Add("ds_data098", "래피드 역률");
            dict.Add("ds_data099", "래피드 당월누적전력량");
        }

        public static string GetColumnName(string code)
        {
            dict.TryGetValue(code, out var name);
            return name;
        }

    }
}
