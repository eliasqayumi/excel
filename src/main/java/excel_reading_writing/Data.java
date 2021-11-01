package excel_reading_writing;

public class Data {
    private String id;
    private String gender;
    private String nationality;
    private String placeOfBirth;
    private String stageId;
    private String sectionId;
    private String topic;
    private String semester;
    private String relation;
    private String raisedHands;
    private String announce;
    private String discussion;
    private String parentAnsweringSurvey;
    private String parentSchoolSatisfaction;
    private String studentAbsenceDay;
    private String classes;

    public Data(String id,String gender, String nationality,String placeOfBirth,String stageId, String sectionId,
                String topic, String semester,String relation, String raisedHands, String announce, String discussion,
                String parentAnsweringSurvey, String parentSchoolSatisfaction, String studentAbsenceDay, String classes) {
        this.id=id;
        this.gender = gender;
        this.nationality = nationality;
        this.placeOfBirth = placeOfBirth;
        this.stageId=stageId;
        this.sectionId = sectionId;
        this.topic = topic;
        this.semester=semester;
        this.relation = relation;
        this.raisedHands = raisedHands;
        this.announce = announce;
        this.discussion = discussion;
        this.parentAnsweringSurvey = parentAnsweringSurvey;
        this.parentSchoolSatisfaction = parentSchoolSatisfaction;
        this.studentAbsenceDay = studentAbsenceDay;
        this.classes = classes;
    }

    public String getStageId() {
        return stageId;
    }

    public void setStageId(String stageId) {
        this.stageId = stageId;
    }

    public String getSemester() {
        return semester;
    }

    public void setSemester(String semester) {
        this.semester = semester;
    }

    public Data() {
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String getNationality() {
        return nationality;
    }

    public void setNationality(String nationality) {
        this.nationality = nationality;
    }

    public String getPlaceOfBirth() {
        return placeOfBirth;
    }

    public void setPlaceOfBirth(String placeOfBirth) {
        this.placeOfBirth = placeOfBirth;
    }

    public String getSectionId() {
        return sectionId;
    }

    public void setSectionId(String sectionId) {
        this.sectionId = sectionId;
    }

    public String getTopic() {
        return topic;
    }

    public void setTopic(String topic) {
        this.topic = topic;
    }

    public String getRelation() {
        return relation;
    }

    public void setRelation(String relation) {
        this.relation = relation;
    }

    public String getRaisedHands() {
        return raisedHands;
    }

    public void setRaisedHands(String raisedHands) {
        this.raisedHands = raisedHands;
    }

    public String getAnnounce() {
        return announce;
    }

    public void setAnnounce(String announce) {
        this.announce = announce;
    }

    public String getDiscussion() {
        return discussion;
    }

    public void setDiscussion(String discussion) {
        this.discussion = discussion;
    }

    public String getParentAnsweringSurvey() {
        return parentAnsweringSurvey;
    }

    public void setParentAnsweringSurvey(String parentAnsweringSurvey) {
        this.parentAnsweringSurvey = parentAnsweringSurvey;
    }

    public String getParentSchoolSatisfaction() {
        return parentSchoolSatisfaction;
    }

    public void setParentSchoolSatisfaction(String parentSchoolSatisfaction) {
        this.parentSchoolSatisfaction = parentSchoolSatisfaction;
    }

    public String getStudentAbsenceDay() {
        return studentAbsenceDay;
    }

    public void setStudentAbsenceDay(String studentAbsenceDay) {
        this.studentAbsenceDay = studentAbsenceDay;
    }

    public String getClasses() {
        return classes;
    }

    public void setClasses(String classes) {
        this.classes = classes;
    }

    @Override
    public String toString() {
        return "Data{" +
                "id='" + id + '\'' +
                ", gender='" + gender + '\'' +
                ", nationality='" + nationality + '\'' +
                ", placeOfBirth='" + placeOfBirth + '\'' +
                ", stageId='" + stageId + '\'' +
                ", selectionId='" + sectionId + '\'' +
                ", topic='" + topic + '\'' +
                ", semester='" + semester + '\'' +
                ", relation='" + relation + '\'' +
                ", raisedHands='" + raisedHands + '\'' +
                ", announce='" + announce + '\'' +
                ", discussion='" + discussion + '\'' +
                ", parentAnsweringSurvey='" + parentAnsweringSurvey + '\'' +
                ", parentSchoolSatisfaction='" + parentSchoolSatisfaction + '\'' +
                ", studentAbsenceDay='" + studentAbsenceDay + '\'' +
                ", classes='" + classes + '\'' +
                '}';
    }
}
