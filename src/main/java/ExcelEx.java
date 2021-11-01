import excel_reading_writing.Data;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelEx {
    private static List<Data> dataList = new ArrayList<>();
    private static List<Data> myDataList = new ArrayList<Data>();
    private static List<String> value;


    public static void main(String[] args) {

        fetchingData();
        setMyDataList();
        changeValuesForGender();
        changeNationality();
        changePlaceOfBirth();
        changeStageID();
        changeSectionId();
        changeTopic();
        changeSemester();
        changeRelation();
        changeParentAnswer();
        changeParentSchoolSatisfaction();
        changeStudentAbsenceDay();
        changeClass();
        showMyDataList();
        announceStandardDeviation();
        raisedHandStandardDeviation();
        discussionStandardDeviation();
        showMyDataList();
        interFemaleAndMale();
        interNationality();
        interParentAnswer();
        interPlaceOfBirth();
        interSemester();
        interRelation();
        interParentSchoolSatisfaction();
        interTopic();
        interStageID();
        interSectionID();
        interClass();
        interStudentAbsentDay();


    }

    private static void announceStandardDeviation() {
        double sum = 0.0;
        double standardDeviation = 0.0;
        double mean = 0.0;
        double res = 0.0;
        double sq = 0.0;
        int n = myDataList.size();
        for (Data data : myDataList)
            sum = sum + Double.parseDouble(data.getAnnounce());
        mean = sum / (n);
        for (Data data : myDataList)
            standardDeviation = standardDeviation + Math.pow((Double.parseDouble(data.getAnnounce()) - mean), 2);
        sq = standardDeviation / n;
        res = Math.sqrt(sq);
        System.out.println("Announce Standard deviation");
        System.out.println("Standard Sapma " + res);
        System.out.println("Ortalama " + mean);
        System.out.println("Veri setimizin sayisi " + n);
        System.out.println("Merkezden 2 standard sapma degeri " + (mean + 2 * res));
        for (Data data : myDataList)
            if (Double.parseDouble(data.getAnnounce()) >= (mean + 2 * res)) {
                System.out.println("removed for Announce column " + data);
                data.setAnnounce(String.valueOf(mean));
            }
        System.out.println("*************************************\n\n");
    }

    private static void raisedHandStandardDeviation() {
        double sum = 0.0;
        double standardDeviation = 0.0;
        double mean = 0.0;
        double res = 0.0;
        double sq = 0.0;
        int n = myDataList.size();
        for (Data data : myDataList)
            sum = sum + Double.parseDouble(data.getRaisedHands());
        mean = sum / (n);
        for (Data data : myDataList)
            standardDeviation = standardDeviation + Math.pow((Double.parseDouble(data.getRaisedHands()) - mean), 2);
        sq = standardDeviation / n;
        res = Math.sqrt(sq);
        System.out.println("Raised hand Standard deviation");
        System.out.println("Standard Sapma " + res);
        System.out.println("Ortalama " + mean);
        System.out.println("Veri setimizin sayisi " + n);
        System.out.println("Merkezden 2 standard sapma degeri " + (mean + 2 * res));
        for (Data data : myDataList)
            if (Double.parseDouble(data.getAnnounce()) >= (mean + 2 * res))
                System.out.println("removed for Announce column " + data);
        System.out.println("*************************************\n\n");

    }

    private static void discussionStandardDeviation() {
        double sum = 0.0;
        double standardDeviation = 0.0;
        double mean = 0.0;
        double res = 0.0;
        double sq = 0.0;
        int n = myDataList.size();
        for (Data data : myDataList)
            sum = sum + Double.parseDouble(data.getDiscussion());
        mean = sum / (n);
        for (Data data : myDataList)
            standardDeviation = standardDeviation + Math.pow((Double.parseDouble(data.getDiscussion()) - mean), 2);
        sq = standardDeviation / n;
        res = Math.sqrt(sq);
        System.out.println("Discussion Standard deviation");
        System.out.println("Standard Sapma " + res);
        System.out.println("Ortalama " + mean);
        System.out.println("Veri setimizin sayisi " + n);
        System.out.println("Merkezden 2 standard sapma degeri " + (mean + 2 * res));
        for (Data data : myDataList)
            if (Double.parseDouble(data.getAnnounce()) >= (mean + 2 * res))
                System.out.println("removed for Announce column " + data);
        System.out.println("*************************************\n\n");

    }

    public static void setMyDataList() {
        for (Data data : dataList)
            if (Double.parseDouble(data.getId()) % 3 == 1)
                myDataList.add(data);
    }

    public static void showMyDataList() {
        for (Data data : myDataList)
            System.out.println(data);
        System.out.println("Veri setimiz " + myDataList.size());
    }

    //
    private static void changeNationality() {
        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList) {
            if (!value.contains(data.getNationality())) {
                value.add(data.getNationality());
                data.setNationality(String.valueOf(i));
                i++;
            } else {
                data.setNationality(String.valueOf(value.indexOf(data.getNationality())));
            }
        }
        for (String string : value)
            System.out.println(string);
        System.out.println("*************************");
    }

    private static void interNationality() {
        double kw = 0, saud = 0, jor = 0, usa = 0, leb = 0, iran = 0, egypt = 0, tunis = 0, moroc = 0, syr = 0, pales = 0, iraq = 0, lybia = 0;
        for (Data data : myDataList) {
            switch (data.getNationality()) {
                case "0":
                    kw++;
                    break;
                case "1":
                    saud++;
                    break;
                case "2":
                    jor++;
                    break;
                case "3":
                    usa++;
                    break;
                case "4":
                    leb++;
                    break;
                case "5":
                    iran++;
                    break;
                case "6":
                    egypt++;
                    break;
                case "7":
                    tunis++;
                    break;
                case "8":
                    moroc++;
                    break;
                case "9":
                    syr++;
                    break;
                case "10":
                    pales++;
                    break;
                case "11":
                    iraq++;
                    break;
                case "12":
                    lybia++;
                    break;
            }
        }
//        System.out.println("KW " + kw + " Saud " + saud + " JOR " + jor + " USA " + usa + " Leb " +
//                leb + " Iran " + iran + " egypt " + egypt + " tunis " + tunis + " moroc " + moroc + " Syr " + syr + " Palas " + pales +
//                " Iraq " + iraq + " Lyba" + lybia);

        double inter = (kw / myDataList.size()) * (-Math.log(kw / myDataList.size())) +
                (saud / myDataList.size()) * (-Math.log(saud / myDataList.size())) +
                (jor / myDataList.size()) * (-Math.log(jor / myDataList.size())) +
                (usa / myDataList.size()) * (-Math.log(usa / myDataList.size())) +
                (leb / myDataList.size()) * (-Math.log(leb / myDataList.size())) +
                (iran / myDataList.size()) * (-Math.log(iran / myDataList.size())) +
                (egypt / myDataList.size()) * (-Math.log(egypt / myDataList.size())) +
                (tunis / myDataList.size()) * (-Math.log(tunis / myDataList.size())) +
                (moroc / myDataList.size()) * (-Math.log(moroc / myDataList.size())) +
                (syr / myDataList.size()) * (-Math.log(syr / myDataList.size())) +
                (pales / myDataList.size()) * (-Math.log(pales / myDataList.size())) +
                (iraq / myDataList.size()) * (-Math.log(iraq / myDataList.size())) +
                (lybia / myDataList.size()) * (-Math.log(lybia / myDataList.size()));

        System.out.println("Natinality Interapy " + (1 - inter));

    }

    // Parent Answer
    private static void changeParentAnswer() {
        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList)
            if (!value.contains(data.getParentAnsweringSurvey())) {
                value.add(data.getParentAnsweringSurvey());
                data.setParentAnsweringSurvey(String.valueOf(i));
                i++;
            } else
                data.setParentAnsweringSurvey(String.valueOf(value.indexOf(data.getParentAnsweringSurvey())));
        for (String string : value)
            System.out.println(string);
        System.out.println("*************************");
    }

    private static void interParentAnswer() {
        double yes = 0;
        double no = 0;
        for (Data data : myDataList)
            switch (data.getParentAnsweringSurvey()) {
                case "0":
                    yes++;
                    break;
                case "1":
                    no++;
                    break;
            }
        double inter = (yes / myDataList.size()) * (-Math.log(yes / myDataList.size())) +
                (no / myDataList.size()) * (-Math.log(no / myDataList.size()));
        System.out.println("Parent Interapy " + (inter));
        System.out.println("Parent information gain " + (1 - inter));
        System.out.println("******************************************\n");
    }

    // Place Of birth
    private static void changePlaceOfBirth() {
        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList)
            if (!value.contains(data.getPlaceOfBirth())) {
                value.add(data.getPlaceOfBirth());
                data.setPlaceOfBirth(String.valueOf(i));
                i++;
            } else
                data.setPlaceOfBirth(String.valueOf(value.indexOf(data.getPlaceOfBirth())));
        for (String string : value)
            System.out.println(string);
        System.out.println("*************************");
    }

    private static void interPlaceOfBirth() {
        double kw = 0, saud = 0, jor = 0, usa = 0, leb = 0, iran = 0, egypt = 0, tunis = 0, moroc = 0, syr = 0, pales = 0, iraq = 0, lybia = 0;
        for (Data data : myDataList) {
            switch (data.getPlaceOfBirth()) {
                case "0":
                    kw++;
                    break;
                case "1":
                    saud++;
                    break;
                case "2":
                    jor++;
                    break;
                case "3":
                    usa++;
                    break;
                case "4":
                    leb++;
                    break;
                case "5":
                    iran++;
                    break;
                case "6":
                    egypt++;
                    break;
                case "7":
                    tunis++;
                    break;
                case "8":
                    moroc++;
                    break;
                case "9":
                    iraq++;
                    break;
                case "10":
                    syr++;
                    break;
                case "11":
                    lybia++;
                    break;
                case "12":
                    pales++;
                    break;
            }
        }
//        System.out.println("KW " + kw + " Saud " + saud + " JOR " + jor + " USA " + usa + " Leb " +
//                leb + " Iran " + iran + " egypt " + egypt + " tunis " + tunis + " moroc " + moroc + " Syr " + syr + " Palas " + pales +
//                " Iraq " + iraq + " Lyba" + lybia);

        double inter = (kw / myDataList.size()) * (-Math.log(kw / myDataList.size())) +
                (saud / myDataList.size()) * (-Math.log(saud / myDataList.size())) +
                (jor / myDataList.size()) * (-Math.log(jor / myDataList.size())) +
                (usa / myDataList.size()) * (-Math.log(usa / myDataList.size())) +
                (leb / myDataList.size()) * (-Math.log(leb / myDataList.size())) +
                (iran / myDataList.size()) * (-Math.log(iran / myDataList.size())) +
                (egypt / myDataList.size()) * (-Math.log(egypt / myDataList.size())) +
                (tunis / myDataList.size()) * (-Math.log(tunis / myDataList.size())) +
                (moroc / myDataList.size()) * (-Math.log(moroc / myDataList.size())) +
                (syr / myDataList.size()) * (-Math.log(syr / myDataList.size())) +
                (pales / myDataList.size()) * (-Math.log(pales / myDataList.size())) +
                (iraq / myDataList.size()) * (-Math.log(iraq / myDataList.size())) +
                (lybia / myDataList.size()) * (-Math.log(lybia / myDataList.size()));

        System.out.println("Place Of Birth Interapy " + (inter));
        System.out.println("Place Of Birth Information gain " + (1 - inter));
        System.out.println("******************************************\n");


    }

    // Semester
    private static void changeSemester() {
        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList)
            if (!value.contains(data.getSemester())) {
                value.add(data.getSemester());
                data.setSemester(String.valueOf(i));
                i++;
            } else
                data.setSemester(String.valueOf(value.indexOf(data.getSemester())));
        for (String string : value)
            System.out.println(string);
        System.out.println("*************************");
    }

    public static void interSemester() {
        double f = 0;
        double s = 0;
        for (Data data : myDataList) {
            if (data.getSemester().equals("0")) {
                f++;
            } else
                s++;
        }
        System.out.println("Fall " + f + "Spring " + s);
        double inter = (f / myDataList.size()) * (-Math.log(f / myDataList.size())) +
                (s / myDataList.size()) * (-Math.log(s / myDataList.size()));
        System.out.println("Semester Interapy " + (inter));
        System.out.println("Semester Information gain " + (1 - inter));
        System.out.println("******************************************\n");
    }

    // Relation
    private static void changeRelation() {
        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList)
            if (!value.contains(data.getRelation())) {
                value.add(data.getRelation());
                data.setRelation(String.valueOf(i));
                i++;
            } else
                data.setRelation(String.valueOf(value.indexOf(data.getRelation())));
        for (String string : value)
            System.out.println(string);
        System.out.println("*************************");
    }

    public static void interRelation() {
        double father = 0;
        double mother = 0;
        for (Data data : myDataList) {
            if (data.getRelation().equals("0")) {
                father++;
            } else
                mother++;
        }
        System.out.println("Father " + father + " Mother " + mother);
        double inter = (father / myDataList.size()) * (-Math.log(father / myDataList.size())) +
                (mother / myDataList.size()) * (-Math.log(mother / myDataList.size()));
        System.out.println("Relation Interopy " + (inter));
        System.out.println("Relation Information gain " + (1 - inter));
        System.out.println("******************************************\n");
    }

    // Stage ID
    private static void changeStageID() {
        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList) {
            if (!value.contains(data.getStageId())) {
                value.add(data.getStageId());
                data.setStageId(String.valueOf(i));
                i++;
            } else {
                data.setStageId(String.valueOf(value.indexOf(data.getStageId())));
            }
        }
        for (String string : value) {
            System.out.println(string);
        }
        System.out.println("***********************");
    }

    public static void interStageID() {
        double lower = 0;
        double middle = 0;
        double high = 0;
        for (Data data : myDataList) {
            switch (data.getStageId()) {
                case "0":
                    lower++;
                    break;
                case "1":
                    middle++;
                    break;
                case "2":
                    high++;
                    break;
            }
        }
        System.out.println("High " + high + " Middle " + middle + " Low" + lower);
        double inter = (lower / myDataList.size()) * (-Math.log(lower / myDataList.size())) +
                (middle / myDataList.size()) * (-Math.log(middle / myDataList.size())) +
                (high / myDataList.size()) * (-Math.log(high / myDataList.size()));
        System.out.println("Stage Id Interopy " + (inter));
        System.out.println("Stage Id Information gain " + (1 - inter));
        System.out.println("******************************************\n");
    }

    //Parent School Satisfaction
    private static void changeParentSchoolSatisfaction() {
        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList)
            if (!value.contains(data.getParentSchoolSatisfaction())) {
                value.add(data.getParentSchoolSatisfaction());
                data.setParentSchoolSatisfaction(String.valueOf(i));
                i++;
            } else
                data.setParentSchoolSatisfaction(String.valueOf(value.indexOf(data.getParentSchoolSatisfaction())));
        for (String string : value)
            System.out.println(string);
        System.out.println("*************************");
    }

    private static void interParentSchoolSatisfaction() {
        double good = 0;
        double bad = 0;
        for (Data data : myDataList) {
            if (data.getParentSchoolSatisfaction().equals("0")) {
                good++;
            } else
                bad++;
        }
        System.out.println("Good " + good + " Bad " + bad);
        double inter = (good / myDataList.size()) * (-Math.log(good / myDataList.size())) +
                (bad / myDataList.size()) * (-Math.log(bad / myDataList.size()));
        System.out.println("Parent School Satisfaction Interopy " + (inter));
        System.out.println("Parent School Satisfaction Information gain " + (1 - inter));
        System.out.println("******************************************\n");
    }

    //Class
    private static void changeClass() {
        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList)
            if (!value.contains(data.getClasses())) {
                value.add(data.getClasses());
                data.setClasses(String.valueOf(i));
                i++;
            } else
                data.setClasses(String.valueOf(value.indexOf(data.getClasses())));
        for (String string : value)
            System.out.println(string);
        System.out.println("*************************");
    }

    public static void interClass() {
        double h = 0;
        double m = 0;
        double l = 0;
        for (Data data : myDataList) {
            switch (data.getClasses()) {
                case "0":
                    m++;
                    break;
                case "1":
                    l++;
                    break;
                case "2":
                    h++;
                    break;
            }
        }
        System.out.println("Hieght " + h + " Middle " + m + " Low" + l);
        double inter = (h / myDataList.size()) * (-Math.log(h / myDataList.size())) +
                (m / myDataList.size()) * (-Math.log(m / myDataList.size())) +
                (l / myDataList.size()) * (-Math.log(l / myDataList.size()));
        System.out.println("Class Interopy " + (inter));
        System.out.println("Class Interopy  Information gain " + (1 - inter));
        System.out.println("******************************************\n");
    }

    // Section
    private static void changeSectionId() {
        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList)
            if (!value.contains(data.getSectionId())) {
                value.add(data.getSectionId());
                data.setSectionId(String.valueOf(i));
                i++;
            } else
                data.setSectionId(String.valueOf(value.indexOf(data.getSectionId())));
        for (String string : value)
            System.out.println(string);
        System.out.println("*************************");
    }

    public static void interSectionID() {
        double A = 0;
        double B = 0;
        double C = 0;
        for (Data data : myDataList) {
            switch (data.getSectionId()) {
                case "0":
                    A++;
                    break;
                case "1":
                    B++;
                    break;
                case "2":
                    C++;
                    break;
            }
        }
        System.out.println("A " + A + " B " + B + " C " + C);
        double inter = (A / myDataList.size()) * (-Math.log(A / myDataList.size())) +
                (B / myDataList.size()) * (-Math.log(B / myDataList.size())) +
                (C / myDataList.size()) * (-Math.log(C / myDataList.size()));
        System.out.println("Section Interopy " + ( inter));
        System.out.println("Section Information gain " + (1 - inter));
        System.out.println("******************************************\n");
    }

    // Topic
    private static void changeTopic() {
        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList)
            if (!value.contains(data.getTopic())) {
                value.add(data.getTopic());
                data.setTopic(String.valueOf(i));
                i++;
            } else
                data.setTopic(String.valueOf(value.indexOf(data.getTopic())));
        for (String string : value)
            System.out.println(string);
        System.out.println("*************************");
    }

    private static void interTopic() {
        double it = 0, math = 0, arabic = 0, english = 0, science = 0, quran = 0, spanish = 0, french = 0, history = 0,
                biology = 0, chemistry = 0, geology = 0;
        for (Data data : myDataList) {
            switch (data.getTopic()) {
                case "0":
                    it++;
                    break;
                case "1":
                    math++;
                    break;
                case "2":
                    arabic++;
                    break;
                case "3":
                    english++;
                    break;
                case "4":
                    science++;
                    break;
                case "5":
                    quran++;
                    break;
                case "6":
                    spanish++;
                    break;
                case "7":
                    french++;
                    break;
                case "8":
                    history++;
                    break;
                case "9":
                    geology++;
                    break;
                case "10":
                    biology++;
                    break;
                case "11":
                    chemistry++;
                    break;
            }
        }
//        System.out.println("KW " + kw + " Saud " + saud + " JOR " + jor + " USA " + usa + " Leb " +
//                leb + " Iran " + iran + " egypt " + egypt + " tunis " + tunis + " moroc " + moroc + " Syr " + syr + " Palas " + pales +
//                " Iraq " + iraq + " Lyba" + lybia);

        double inter = (it / myDataList.size()) * (-Math.log(it / myDataList.size())) +
                (math / myDataList.size()) * (-Math.log(math / myDataList.size())) +
                (arabic / myDataList.size()) * (-Math.log(arabic / myDataList.size())) +
                (english / myDataList.size()) * (-Math.log(english / myDataList.size())) +
                (science / myDataList.size()) * (-Math.log(science / myDataList.size())) +
                (quran / myDataList.size()) * (-Math.log(quran / myDataList.size())) +
                (spanish / myDataList.size()) * (-Math.log(spanish / myDataList.size())) +
                (french / myDataList.size()) * (-Math.log(french / myDataList.size())) +
                (history / myDataList.size()) * (-Math.log(history / myDataList.size())) +
                (biology / myDataList.size()) * (-Math.log(biology / myDataList.size())) +
                (chemistry / myDataList.size()) * (-Math.log(chemistry / myDataList.size())) +
                (geology / myDataList.size()) * (-Math.log(geology / myDataList.size()));

        System.out.println("Topic Interapy " + (inter));
        System.out.println("Topic Information gain " + (1 - inter));
        System.out.println("******************************************\n");
    }

    //    Student Absence day
    private static void changeStudentAbsenceDay() {

        int i = 0;
        value = new ArrayList<>();
        for (Data data : myDataList)
            if (!value.contains(data.getStudentAbsenceDay())) {
                value.add(data.getStudentAbsenceDay());
                data.setStudentAbsenceDay(String.valueOf(i));
                i++;
            } else
                data.setStudentAbsenceDay(String.valueOf(value.indexOf(data.getStudentAbsenceDay())));
        for (String string : value)
            System.out.println(string);
        System.out.println("*************************");
    }

    private static void interStudentAbsentDay() {
        double above = 0;
        double under = 0;
        for (Data data : myDataList) {
            if (data.getStudentAbsenceDay().equals("0")) {
                under++;
            } else
                above++;
        }
        System.out.println("Under-7 " + under + " Above-7 " + above);
        double inter = (under / myDataList.size()) * (-Math.log(under / myDataList.size())) +
                (above / myDataList.size()) * (-Math.log(above / myDataList.size()));
        System.out.println("Student Absence day Interopy " + ( inter));
        System.out.println("Student Absence day Information gain " + (1 - inter));
        System.out.println("******************************************\n");
    }
    //Gender
    // Male = 0      Female = 1
    private static void changeValuesForGender() {
        for (Data data : myDataList)
            if (data.getGender().equals("M"))
                data.setGender("0");
            else
                data.setGender("1");
    }

    public static void interFemaleAndMale() {
        double female = 0;
        double male = 0;
        for (Data data : myDataList) {
            if (data.getGender().equals("1")) {
                female++;
            } else
                male++;
        }
        System.out.println("Male " + male + "Female " + female);
        System.out.println("Veri sayisi" + myDataList.size());
        double inter = (female / myDataList.size()) * (-Math.log(female / myDataList.size())) +
                (male / myDataList.size()) * (-Math.log(male / myDataList.size()));
        System.out.println("Gender Interapy " + (inter));
        System.out.println("Gender Information gain " + (1 - inter));
        System.out.println("******************************************\n");
    }


    private static void fetchingData() {
        String excelPath = "/Users/qayumi/Desktop/All_Codes/FullWebAppStore/excel/Data/edu_data.xlsx";
        try {
            FileInputStream inputStream = new FileInputStream(excelPath);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");
            int rowNum = sheet.getLastRowNum();
            int columnNum = sheet.getRow(0).getLastCellNum();
            for (int i = 1; i < rowNum; i++) {
                dataList.add(new Data(
                        String.valueOf(sheet.getRow(i).getCell(0).toString()),
                        String.valueOf(sheet.getRow(i).getCell(1).toString()),
                        String.valueOf(sheet.getRow(i).getCell(2).toString()),
                        String.valueOf(sheet.getRow(i).getCell(3).toString()),
                        String.valueOf(sheet.getRow(i).getCell(4).toString()),
                        String.valueOf(sheet.getRow(i).getCell(5).toString()),
                        String.valueOf(sheet.getRow(i).getCell(6).toString()),
                        String.valueOf(sheet.getRow(i).getCell(7).toString()),
                        String.valueOf(sheet.getRow(i).getCell(8).toString()),
                        String.valueOf(sheet.getRow(i).getCell(9).toString()),
                        String.valueOf(sheet.getRow(i).getCell(10).toString()),
                        String.valueOf(sheet.getRow(i).getCell(11).toString()),
                        String.valueOf(sheet.getRow(i).getCell(12).toString()),
                        String.valueOf(sheet.getRow(i).getCell(13).toString()),
                        String.valueOf(sheet.getRow(i).getCell(14).toString()),
                        String.valueOf(sheet.getRow(i).getCell(15).toString())
                ));

            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
