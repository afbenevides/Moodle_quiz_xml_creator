import xlwings as xw
import unicodedata
import xml.etree.ElementTree as ET


class xlsx_opener():
    def __init__(self, data_file_name):
        self.question_data_list = []
        try:
            xw.apps
            print(xw.apps)
        except ValueError:
            print("Oops!  That was no valid number.  Try again...")

        try:
            xw.books.open(data_file_name)
        except ValueError:
            print("incapable d'ouvrir le ichier excel, p-e ouverture en double?")
        sheet_name = 'ListeQuestions'
        print(sheet_name)
        xw.sheets[sheet_name].activate()
        # lwr_r_cell = xw.cells.last_cell  # lower right cell
        self.last_line = xw.Range('D3').end('down').row

        # Read all data from file
        for line in range(4, self.last_line + 1):
            print("one line " + str(line))
            self.read_question_items(line)
        # print(self.question_list)

    def read_question_items(self, line_number):
        question = self.read_cell('D' + str(line_number))
        type = self.read_cell('E' + str(line_number))
        choix1 = self.read_cell('F' + str(line_number))
        choix2 = self.read_cell('G' + str(line_number))
        choix3 = self.read_cell('H' + str(line_number))
        choix4 = self.read_cell('I' + str(line_number))
        choix5 = self.read_cell('J' + str(line_number))
        reponse1 = self.read_cell('K' + str(line_number))
        reponse2 = self.read_cell('L' + str(line_number))
        reponse3 = self.read_cell('M' + str(line_number))
        reponse4 = self.read_cell('N' + str(line_number))
        reponse5 = self.read_cell('O' + str(line_number))
        parameter_list = [question, type, choix1, choix2, choix3, choix4, choix5, reponse1, reponse2, reponse3,
                          reponse4,
                          reponse5]
        print(parameter_list)
        self.question_data_list.append(parameter_list)

    def read_cell(self, cell_number):
        value = xw.Range(cell_number).value
        if isinstance(value, str):
            return unicodedata.normalize("NFKD", xw.Range(cell_number).value)
        else:
            return value


class quiz_xml():
    def __init__(self, question_data_list):
        self.quiz = ET.Element('quiz')
        comment = ET.Comment(' === Some Comment === ')
        self.quiz.insert(1, comment)  # 1 is the index where comment is inserted
        for question_data in question_data_list:
            if question_data[1] == 'Choix multiple simple':
                quiz_question_Choix_multiple_simple(self.quiz, question_data)
            elif question_data[1] == 'Choix multiple checkbox':
                quiz_question_Choix_multiple_checkbox(self.quiz, question_data)
            elif question_data[1] == 'Vrai ou Faux':
                quiz_question_true_false(self.quiz, question_data)
            else:
                print("Ce type de question n'est pas encore suporté")
        self.write_quiz()

    def add_question(self, question):
        ET.SubElement(self.quiz, question)

    def write_quiz(self):
        print(self.quiz)
        q_data = ET.tostring(self.quiz)
        output_file = open("output_quiz.xml", "w")
        replacements = {
            "&gt;": ">",
            "&lt;": "<",
            "&#233;": "é"
        }
        q_data = q_data.decode("utf-8")
        for char in replacements:
            q_data = q_data.replace(char, replacements[char])
        output_file.write(q_data)


class quiz_question():
    def __init__(self, quiz, parameter_list):
        self.question = ET.SubElement(quiz, 'question')

        # name
        name = ET.SubElement(self.question, 'name')
        name_text = ET.SubElement(name, 'text')
        name_text.text = parameter_list[0]

        # questiontext
        questiontext = ET.SubElement(self.question, 'questiontext')
        questiontext.set('format', 'html')
        questiontext_text = ET.SubElement(questiontext, 'text')
        questiontext_text.text = "<![CDATA[<p>" + parameter_list[0] + "?</p>]]>"

        # generalfeedback
        generalfeedback = ET.SubElement(self.question, 'generalfeedback')
        generalfeedback.set('format', 'html')
        generalfeedback_text = ET.SubElement(generalfeedback, 'text')
        generalfeedback_text.text = ""

        # defaultgrade
        defaultgrade = ET.SubElement(self.question, 'defaultgrade')
        defaultgrade.text = "1.0000000"

        # penalty
        penalty = ET.SubElement(self.question, 'penalty')
        penalty.text = "1.0000000"

        # hidden
        hidden = ET.SubElement(self.question, 'hidden')
        hidden.text = "0"

        # idnumber
        idnumber = ET.SubElement(self.question, 'idnumber')


        # good_answer_count
        good_answer_count = 0
        for answer_reponse in parameter_list[7:11]:
            if answer_reponse is None or answer_reponse == 0:
                pass
            else:
                good_answer_count += 1
        fraction = 100 / good_answer_count

        # for answer_no in range(2,1+good_answer_count):
        #     answer_element = parameter_list[answer_no]
        #     if answer_element is not None or answer_element != 0:
        #         answer = ET.SubElement(self.question, 'answer')
        #         if parameter_list[answer_no+5]==1:
        #             answer.set('fraction', str(fraction))
        #         else:
        #             answer.set('fraction', str(0))
        #         answer.set('format', 'html')
        #         answer_text = ET.SubElement(answer, 'text')
        #         answer_text.text = "<![CDATA[<p>"+ answer_element + "</p>]]>"
        #
        #         # generalfeedback
        #         answer_feedback = ET.SubElement(answer, 'feedback')
        #         answer_feedback.set('format', 'html')
        #         answer_feedback_text = ET.SubElement(answer_feedback, 'text')
        #         answer_feedback_text.text = ""
        #
        if parameter_list[2] is not None and parameter_list[2] != 0:
            answer1 = ET.SubElement(self.question, 'answer')
            if parameter_list[7] == 1:
                answer1.set('fraction', str(fraction))
            else:
                answer1.set('fraction', str(0))
            answer1.set('format', 'html')
            answer1_text = ET.SubElement(answer1, 'text')
            answer1_text.text = "<![CDATA[<p>" + parameter_list[2] + "</p>]]>"

            # generalfeedback
            answer1_feedback = ET.SubElement(answer1, 'feedback')
            answer1_feedback.set('format', 'html')
            answer1_feedback_text = ET.SubElement(answer1_feedback, 'text')
            answer1_feedback_text.text = ""

        if parameter_list[3] is not None and parameter_list[3] != 0:
            answer2 = ET.SubElement(self.question, 'answer')
            if parameter_list[8] == 1:
                answer2.set('fraction', str(fraction))
            else:
                answer2.set('fraction', str(0))
            answer2.set('format', 'html')
            answer2_text = ET.SubElement(answer2, 'text')
            answer2_text.text = "<![CDATA[<p>" + parameter_list[3] + "</p>]]>"

            # generalfeedback
            answer2_feedback = ET.SubElement(answer2, 'feedback')
            answer2_feedback.set('format', 'html')
            answer2_feedback_text = ET.SubElement(answer2_feedback, 'text')
            answer2_feedback_text.text = ""

        if parameter_list[4] is not None and parameter_list[4] != 0:
            answer3 = ET.SubElement(self.question, 'answer')
            if parameter_list[9] == 1:
                answer3.set('fraction', str(fraction))
            else:
                answer3.set('fraction', str(0))
            answer3.set('format', 'html')
            answer3_text = ET.SubElement(answer3, 'text')
            answer3_text.text = "<![CDATA[<p>" + parameter_list[4] + "</p>]]>"

            # generalfeedback
            answer3_feedback = ET.SubElement(answer3, 'feedback')
            answer3_feedback.set('format', 'html')
            answer3_feedback_text = ET.SubElement(answer3_feedback, 'text')
            answer3_feedback_text.text = ""
        
        if parameter_list[5] is not None and parameter_list[5] != 0:
            answer4 = ET.SubElement(self.question, 'answer')
            if parameter_list[10] == 1:
                answer4.set('fraction', str(fraction))
            else:
                answer4.set('fraction', str(0))
            answer4.set('format', 'html')
            answer4_text = ET.SubElement(answer4, 'text')
            answer4_text.text = "<![CDATA[<p>" + parameter_list[5] + "</p>]]>"

            # generalfeedback
            answer4_feedback = ET.SubElement(answer4, 'feedback')
            answer4_feedback.set('format', 'html')
            answer4_feedback_text = ET.SubElement(answer4_feedback, 'text')
            answer4_feedback_text.text = ""
            
        if parameter_list[6] is not None and parameter_list[6] != 0:
            answer5 = ET.SubElement(self.question, 'answer')
            if parameter_list[11] == 1:
                answer5.set('fraction', str(fraction))
            else:
                answer5.set('fraction', str(0))
            answer5.set('format', 'html')
            answer5_text = ET.SubElement(answer5, 'text')
            answer5_text.text = "<![CDATA[<p>" + parameter_list[6] + "</p>]]>"

            # generalfeedback
            answer5_feedback = ET.SubElement(answer5, 'feedback')
            answer5_feedback.set('format', 'html')
            answer5_feedback_text = ET.SubElement(answer5_feedback, 'text')
            answer5_feedback_text.text = ""





class quiz_question_Choix_multiple_checkbox(quiz_question):
    def __init__(self, quiz, parameter_list):
        super().__init__(quiz, parameter_list)
        self.question.set('type', 'multichoice')

        # single
        self.single = ET.SubElement(self.question, 'single')
        self.single.text = "false"

        # shuffleanswers
        shuffleanswers = ET.SubElement(self.question, 'shuffleanswers')
        shuffleanswers.text = "true"

        # answernumbering
        answernumbering = ET.SubElement(self.question, 'answernumbering')
        answernumbering.text = "abc"

        # correctfeedback
        correctfeedback = ET.SubElement(self.question, 'correctfeedback')
        correctfeedback.set('format', 'html')
        correctfeedback_text = ET.SubElement(correctfeedback, 'text')
        correctfeedback_text.text = "<![CDATA[<p>Votre réponse est correcte.</p>]]>"

        # partiallycorrectfeedback
        partiallycorrectfeedback = ET.SubElement(self.question, 'partiallycorrectfeedback')
        partiallycorrectfeedback.set('format', 'html')
        partiallycorrectfeedback_text = ET.SubElement(partiallycorrectfeedback, 'text')
        partiallycorrectfeedback_text.text = "<![CDATA[<p>Votre réponse est partiellement correcte.</p>]]>"

        # incorrectfeedback
        incorrectfeedback = ET.SubElement(self.question, 'incorrectfeedback')
        incorrectfeedback.set('format', 'html')
        incorrectfeedback_text = ET.SubElement(incorrectfeedback, 'text')
        incorrectfeedback_text.text = "<![CDATA[<p>Votre réponse est partiellement correcte.</p>]]>"

        # shownumcorrect
        shownumcorrect = ET.SubElement(self.question, 'shownumcorrect')


class quiz_question_Choix_multiple_simple(quiz_question_Choix_multiple_checkbox):
    def __init__(self, quiz, parameter_list):
        super().__init__(quiz, parameter_list)
        self.question.set('type', 'multichoice')
        self.single.text = "true"


class quiz_question_true_false(quiz_question):
    def __init__(self, quiz, parameter_list):
        super().__init__(quiz, parameter_list)
        self.question.set('type', 'truefalse')


class quiz():
    pass


if __name__ == '__main__':
    fichier_excel = xlsx_opener("ListeQuestions.xlsx")
    quiz_xml(fichier_excel.question_data_list)
