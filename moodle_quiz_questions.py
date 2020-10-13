import xlwings as xw
from xlwings.constants import DeleteShiftDirection

import unicodedata
import xml.etree.ElementTree as ET
import copy


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
            print("incapable d'ouvrir le fichier excel, p-e ouverture en double?")
        sheet_name = 'ListeQuestions'
        print(sheet_name)
        xw.sheets[sheet_name].activate()
        # lwr_r_cell = xw.cells.last_cell  # lower right cell
        self.last_line = xw.Range('D3').end('down').row

        # Read all data from file
        for line in range(3, self.last_line + 1):
            print("one line " + str(line))
            self.read_question_items(line)
        # print(self.question_list)
        self.question_data_list.sort(key=self.take_last)

    def take_last(self, elem):
        return elem[-1]

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
        category_num = self.read_cell('A' + str(line_number))
        if category_num is None:
            category_num = "9999-0"

        parameter_list = [question, type, choix1, choix2, choix3, choix4, choix5, reponse1, reponse2, reponse3,
                          reponse4, reponse5, category_num]
        print(parameter_list)
        self.question_data_list.append(parameter_list)

    def read_cell(self, cell_number):
        value = xw.Range(cell_number).value
        if isinstance(value, str):
            return unicodedata.normalize("NFKD", xw.Range(cell_number).value)
        else:
            return value


class statistiques():
    def __init__(self, data_file_name, question_data):
        sheet_name = 'Ponderation'
        print(sheet_name)
        xw.sheets[sheet_name].activate()
        # lwr_r_cell = xw.cells.last_cell  # lower right cell
        if xw.Range('A3').value:
            self.last_line = xw.Range('A2').end('down').row
            xw.Range(str(2) + ':' + str(self.last_line+1)).delete()

        #Trouver tous les lignes à traiter

        my_list_name=[]
        my_list_id=[]
        id_list=[]
        for each in question_data:
            id_list.append(each[-1])

        for each in question_data:
            if each[-1][-1] == '0':
                my_list_name.append(each[0].split('/')[-1])
                my_list_id.append(each[-1])

        print(id_list)
        print(my_list_name)
        print(my_list_id)

        #my_unique_list_name = self.unique_id_list(my_list_name)
        #my_unique_list_id = self.unique_id_list(my_list_id)
        #print(my_unique_list_name)
        #print(my_unique_list_id)
        module_qty = 0
        module_qty_sum = len(id_list) - len(my_list_id)
        print(module_qty_sum)

        for increment in range(0,len(my_list_name)):
            xw.Range('A'+str(2+increment)).value = my_list_id[increment]
            xw.Range('B'+str(2+increment)).value = my_list_name[increment]

            # calcul quantité question dans section
            count = 0


            if my_list_id[increment][-1] == '0' and len(my_list_id[increment]) == 5:
                compare_string = my_list_id[increment][0:4] + '1'
                for string in id_list:
                    if compare_string == string:
                        count += 1
                if module_qty != 0:
                    xw.Range('C' + str(2 + increment)).value = count / module_qty_sum
                    xw.Range('G' + str(2 + increment)).value = count / module_qty
                else:
                    xw.Range('C' + str(2 + increment)).value = 0
                    xw.Range('G' + str(2 + increment)).value = 0

            elif my_list_id[increment][-1]== '0' and len(my_list_id[increment]) == 3:
                for string in id_list:
                    if my_list_id[increment][0] == string[0] and string[-1] != '0':
                        count += 1
                xw.Range('A' + str(2 + increment) + ':F' + str(2 + increment)).color = (245, 194, 67)
                module_qty = copy.copy(count)
                #module_qty_sum += module_qty
                if module_qty_sum != 0:
                    xw.Range('C' + str(2 + increment)).value = count / module_qty_sum
                    xw.Range('G' + str(2 + increment)).value = count / module_qty_sum
                else:
                    xw.Range('C' + str(2 + increment)).value = 0
                    xw.Range('G' + str(2 + increment)).value = 0



            xw.Range('F'+str(2+increment)).value = count
            xw.Range('C'+str(2+increment)).number_format = "0.00%"
            xw.Range('G'+str(2+increment)).number_format = "0.00%"





    def unique_id_list(self, list_to_manage):
        return list(dict.fromkeys(list_to_manage))


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
            elif question_data[1] == 'Numerique':
                quiz_question_numerical(self.quiz, question_data)
            elif question_data[1] == 'Reponse courte':
                quiz_question_short_answer(self.quiz, question_data)
            elif question_data[1] == 'Categories':
                quiz_question_categories(self.quiz, question_data)
            else:
                print("Le type de question == " + question_data[1] + " == n'est pas encore supporté")
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
            "&#233;": "é",
            "e&#769;": "é",
            "e&#768;": "è"
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
        for answer_reponse in parameter_list[7:12]:
            if answer_reponse is None or answer_reponse == 0:
                pass
            else:
                good_answer_count += 1

        if parameter_list[1] == 'Reponse courte':
            fraction = 100
        else:
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
        if parameter_list[2] is not None or parameter_list[1] == 'Reponse courte' and parameter_list[7] is not None:
            self.answer1 = ET.SubElement(self.question, 'answer')
            if parameter_list[7] == 1 or parameter_list[1] == 'Numerique' or parameter_list[1] == 'Reponse courte':
                self.answer1.set('fraction', str(fraction))
            else:
                self.answer1.set('fraction', str(0))
            if parameter_list[1] == 'Numerique' or parameter_list[1] == 'Reponse courte':
                self.answer1.set('format', 'moodle_auto_format')
            else:
                self.answer1.set('format', 'html')
            self.answer1_text = ET.SubElement(self.answer1, 'text')
            if parameter_list[1] == 'Reponse courte':
                self.answer1_text.text = str(parameter_list[7])
            elif parameter_list[1] == 'Vrai ou Faux':
                self.answer1_text.text = str(parameter_list[2])
            else:
                self.answer1_text.text = "<![CDATA[<p>" + str(parameter_list[2]) + "</p>]]>"

            # generalfeedback
            answer1_feedback = ET.SubElement(self.answer1, 'feedback')
            answer1_feedback.set('format', 'html')
            answer1_feedback_text = ET.SubElement(answer1_feedback, 'text')
            answer1_feedback_text.text = ""

        if parameter_list[3] is not None or parameter_list[1] == 'Reponse courte' and parameter_list[8] is not None:
            answer2 = ET.SubElement(self.question, 'answer')
            if parameter_list[8] == 1 or parameter_list[1] == 'Reponse courte':
                answer2.set('fraction', str(fraction))
            else:
                answer2.set('fraction', str(0))
            if parameter_list[1] == 'Numerique' or parameter_list[1] == 'Reponse courte':
                answer2.set('format', 'moodle_auto_format')
            else:
                answer2.set('format', 'html')
            answer2.set('format', 'html')
            answer2_text = ET.SubElement(answer2, 'text')

            if parameter_list[1] == 'Reponse courte':
                answer2_text.text = str(parameter_list[8])
            elif parameter_list[1] == 'Vrai ou Faux':
                answer2_text.text = str(parameter_list[3])
            else:
                answer2_text.text = "<![CDATA[<p>" + parameter_list[3] + "</p>]]>"

            # generalfeedback
            answer2_feedback = ET.SubElement(answer2, 'feedback')
            answer2_feedback.set('format', 'html')
            answer2_feedback_text = ET.SubElement(answer2_feedback, 'text')
            answer2_feedback_text.text = ""

        if parameter_list[4] is not None or parameter_list[1] == 'Reponse courte' and parameter_list[9] is not None:
            answer3 = ET.SubElement(self.question, 'answer')
            if parameter_list[9] == 1 or parameter_list[1] == 'Reponse courte':
                answer3.set('fraction', str(fraction))
            else:
                answer3.set('fraction', str(0))
            if parameter_list[1] == 'Numerique' or parameter_list[1] == 'Reponse courte':
                answer3.set('format', 'moodle_auto_format')
            else:
                answer3.set('format', 'html')
            answer3.set('format', 'html')

            answer3_text = ET.SubElement(answer3, 'text')
            if parameter_list[1] == 'Reponse courte':
                answer3_text.text = str(parameter_list[9])
            else:
                answer3_text.text = "<![CDATA[<p>" + parameter_list[4] + "</p>]]>"

            # generalfeedback
            answer3_feedback = ET.SubElement(answer3, 'feedback')
            answer3_feedback.set('format', 'html')
            answer3_feedback_text = ET.SubElement(answer3_feedback, 'text')
            answer3_feedback_text.text = ""

        if parameter_list[5] is not None or parameter_list[1] == 'Reponse courte' and parameter_list[10] is not None:
            answer4 = ET.SubElement(self.question, 'answer')
            if parameter_list[10] == 1 or parameter_list[1] == 'Reponse courte':
                answer4.set('fraction', str(fraction))
            else:
                answer4.set('fraction', str(0))
            if parameter_list[1] == 'Numerique' or parameter_list[1] == 'Reponse courte':
                answer4.set('format', 'moodle_auto_format')
            else:
                answer4.set('format', 'html')
            answer4.set('format', 'html')
            answer4_text = ET.SubElement(answer4, 'text')
            if parameter_list[1] == 'Reponse courte':
                answer4_text.text = str(parameter_list[10])
            else:
                answer4_text.text = "<![CDATA[<p>" + parameter_list[5] + "</p>]]>"

            # generalfeedback
            answer4_feedback = ET.SubElement(answer4, 'feedback')
            answer4_feedback.set('format', 'html')
            answer4_feedback_text = ET.SubElement(answer4_feedback, 'text')
            answer4_feedback_text.text = ""

        if parameter_list[6] is not None or parameter_list[1] == 'Reponse courte' and parameter_list[11] is not None:
            answer5 = ET.SubElement(self.question, 'answer')
            if parameter_list[11] == 1 or parameter_list[1] == 'Reponse courte':
                answer5.set('fraction', str(fraction))
            else:
                answer5.set('fraction', str(0))
            if parameter_list[1] == 'Numerique' or parameter_list[1] == 'Reponse courte':
                answer5.set('format', 'moodle_auto_format')
            else:
                answer5.set('format', 'html')
            answer5.set('format', 'html')
            answer5_text = ET.SubElement(answer5, 'text')
            if parameter_list[1] == 'Reponse courte':
                answer5_text.text = str(parameter_list[11])
            else:
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


class quiz_question_short_answer(quiz_question):
    def __init__(self, quiz, parameter_list):
        super().__init__(quiz, parameter_list)
        self.question.set('type', 'shortanswer')
        usecase = ET.SubElement(self.question, 'usecase')
        usecase.text = "0"


class quiz_question_numerical(quiz_question):
    def __init__(self, quiz, parameter_list):
        super().__init__(quiz, parameter_list)
        self.question.set('type', 'numerical')
        self.answer1_text.text = str(parameter_list[7])
        answer1_tolerance = ET.SubElement(self.answer1, 'tolerance')
        answer1_tolerance.text = str(parameter_list[2])
        unitgradingtype = ET.SubElement(self.question, 'unitgradingtype')
        unitgradingtype.text = "0"
        unitpenalty = ET.SubElement(self.question, 'unitpenalty')
        unitpenalty.text = "0.1000000"
        showunits = ET.SubElement(self.question, 'showunits')
        showunits.text = "3"
        unitsleft = ET.SubElement(self.question, 'unitsleft')
        unitsleft.text = "0"


class quiz_question_categories(quiz_question):
    def __init__(self, quiz, parameter_list):
        self.question = ET.SubElement(quiz, 'question')
        self.question.set('type', 'category')

        # categories
        category = ET.SubElement(self.question, 'category')
        category_text = ET.SubElement(category, 'text')
        category_text.text = "$module$/top/" + str(parameter_list[0])

        # info
        info = ET.SubElement(self.question, 'info')
        info.set('format', 'html')
        info_text = ET.SubElement(info, 'text')
        info_text.text = "<![CDATA[<p>" + str(parameter_list[0]) + "</p>]]>"

        # idnumber
        idnumber = ET.SubElement(self.question, 'idnumber')


if __name__ == '__main__':
    fichier_excel = xlsx_opener("ListeQuestions.xlsx")
    quiz_xml(fichier_excel.question_data_list)
    statistiques("ListeQuestions.xlsx", fichier_excel.question_data_list)
