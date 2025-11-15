"""
MAKERERE UNIVERSITY - COURSEWORK 2
Advanced Calculator System with Student Authentication
Question 4 and 5 Implementation
"""

import math
import os
from datetime import datetime
from docx import Document
from docx.shared import Inches

class AdvancedCalculator:
    """Advanced calculator with arithmetic, scientific functions and history tracking"""
    
    def __init__(self):
        self.calculation_history = []
        self.calculator_document = Document()
        self._setup_calculator_document()
        self.operation_functions = {
            '+': self.add_numbers,
            '-': self.subtract_numbers,
            '*': self.multiply_numbers,
            '/': self.divide_numbers,
            '^': self.calculate_power,
            'sqrt': self.calculate_square_root,
            'sin': self.calculate_sine,
            'cos': self.calculate_cosine,
            'tan': self.calculate_tangent
        }
    
    def _setup_calculator_document(self):
        """Setup the calculator Word document"""
        self.calculator_document.add_heading('MAKERERE UNIVERSITY - ADVANCED CALCULATOR SYSTEM', 0)
        self.calculator_document.add_heading('Mathematical Operations and Calculation History', 1)
        self.calculator_document.add_paragraph(f"Created on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.calculator_document.add_paragraph()
    
    def add_numbers(self, number1, number2):
        """Add two numbers"""
        return number1 + number2
    
    def subtract_numbers(self, number1, number2):
        """Subtract two numbers"""
        return number1 - number2
    
    def multiply_numbers(self, number1, number2):
        """Multiply two numbers"""
        return number1 * number2
    
    def divide_numbers(self, number1, number2):
        """Divide two numbers with zero division check"""
        if number2 == 0:
            raise ValueError("Cannot divide by zero")
        return number1 / number2
    
    def calculate_power(self, base, exponent):
        """Calculate power of a number"""
        return base ** exponent
    
    def calculate_square_root(self, number):
        """Calculate square root with validation"""
        if number < 0:
            raise ValueError("Cannot calculate square root of negative number")
        return math.sqrt(number)
    
    def calculate_sine(self, angle_degrees):
        """Calculate sine of angle in degrees"""
        return math.sin(math.radians(angle_degrees))
    
    def calculate_cosine(self, angle_degrees):
        """Calculate cosine of angle in degrees"""
        return math.cos(math.radians(angle_degrees))
    
    def calculate_tangent(self, angle_degrees):
        """Calculate tangent of angle in degrees"""
        return math.tan(math.radians(angle_degrees))
    
    def execute_operation(self, operation, number1, number2=None):
        """
        Execute the requested operation with comprehensive error handling
        
        Args:
            operation (str): Operation to perform
            number1 (float): First number
            number2 (float, optional): Second number for binary operations
            
        Returns:
            float or str: Result or error message
        """
        try:
            if operation in ['sqrt', 'sin', 'cos', 'tan']:
                result = self.operation_functions[operation](number1)
                operation_description = f"{operation}({number1}) = {result:.4f}"
            else:
                result = self.operation_functions[operation](number1, number2)
                operation_description = f"{number1} {operation} {number2} = {result:.4f}"
            
            # Add to history and document
            history_entry = {
                'operation': operation_description,
                'timestamp': datetime.now().isoformat()
            }
            self.calculation_history.append(history_entry)
            
            # Write to document
            self.calculator_document.add_paragraph(operation_description)
            
            return result
        except KeyError:
            error_message = "Error: Invalid operation selected"
            self.calculator_document.add_paragraph(error_message)
            return error_message
        except ValueError as value_error:
            error_message = f"Error: {str(value_error)}"
            self.calculator_document.add_paragraph(error_message)
            return error_message
        except Exception as general_error:
            error_message = f"Error: {str(general_error)}"
            self.calculator_document.add_paragraph(error_message)
            return error_message
    
    def display_calculation_history(self):
        """Display calculation history in formatted table"""
        self.calculator_document.add_heading('Calculation History', 2)
        
        if not self.calculation_history:
            self.calculator_document.add_paragraph("No calculations performed yet.")
            return
        
        history_table = self.calculator_document.add_table(rows=1, cols=3)
        history_table.style = 'Light Grid Accent 1'
        
        # Add table headers
        header_cells = history_table.rows[0].cells
        header_cells[0].text = 'No.'
        header_cells[1].text = 'Operation'
        header_cells[2].text = 'Timestamp'
        
        # Add history entries
        for index, entry in enumerate(self.calculation_history[-10:], 1):
            row_cells = history_table.add_row().cells
            row_cells[0].text = str(index)
            row_cells[1].text = entry['operation']
            row_cells[2].text = datetime.fromisoformat(entry['timestamp']).strftime('%H:%M:%S')
    
    def save_calculator_document(self):
        """
        Save the calculator Word document
        
        Returns:
            str: Filename of saved document
        """
        try:
            filename = f"Calculator_History_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            self.calculator_document.save(filename)
            print(f"Calculator document saved: {filename}")
            return filename
        except Exception as error:
            print(f"Error saving calculator document: {error}")
            return None
    
    def clear_calculation_history(self):
        """Clear calculation history"""
        self.calculation_history = []
        self.calculator_document.add_paragraph("Calculation history cleared successfully")
        print("Calculation history cleared successfully")

class StudentAuthenticationSystem:
    """Secure student authentication and registration system"""
    
    def __init__(self):
        self.authentication_document = Document()
        self._setup_authentication_document()
        self.registered_users = {}
    
    def _setup_authentication_document(self):
        """Setup the authentication Word document"""
        self.authentication_document.add_heading('MAKERERE UNIVERSITY - STUDENT AUTHENTICATION SYSTEM', 0)
        self.authentication_document.add_heading('User Registration and Login Security Records', 1)
        self.authentication_document.add_paragraph(f"Created on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.authentication_document.add_paragraph()
    
    def register_new_student(self, username, password, registration_number, full_name, student_number):
        """
        Register a new student with validation
        
        Args:
            username (str): Unique username
            password (str): User password
            registration_number (str): University registration number
            full_name (str): Student's full name
            student_number (str): Student identification number
            
        Returns:
            tuple: (success, message)
        """
        if username in self.registered_users:
            self.authentication_document.add_paragraph(f"Registration failed: Username '{username}' already exists")
            return False, "Username already exists"
        
        # Store user data (in production, password should be hashed)
        self.registered_users[username] = {
            'password': password,
            'registration_number': registration_number,
            'full_name': full_name,
            'student_number': student_number,
            'registration_date': datetime.now().isoformat()
        }
        
        # Record registration in document
        self.authentication_document.add_heading(f'New Student Registration: {username}', 3)
        self.authentication_document.add_paragraph(f"Full Name: {full_name}")
        self.authentication_document.add_paragraph(f"Registration Number: {registration_number}")
        self.authentication_document.add_paragraph(f"Student Number: {student_number}")
        self.authentication_document.add_paragraph(f"Registration Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.authentication_document.add_paragraph()
        
        return True, "Registration successful"
    
    def authenticate_student(self, username, password):
        """
        Authenticate student credentials
        
        Args:
            username (str): Student username
            password (str): Student password
            
        Returns:
            tuple: (success, user_data or error_message)
        """
        self.authentication_document.add_paragraph(f"Login attempt for username: {username}")
        
        if username in self.registered_users and self.registered_users[username]['password'] == password:
            user_data = self.registered_users[username]
            self.authentication_document.add_paragraph(f"✓ Successful login for {user_data['full_name']}")
            self.authentication_document.add_paragraph()
            return True, user_data
        
        self.authentication_document.add_paragraph("✗ Login failed: Invalid credentials")
        self.authentication_document.add_paragraph()
        return False, "Invalid credentials"
    
    def save_authentication_document(self):
        """
        Save the authentication Word document
        
        Returns:
            str: Filename of saved document
        """
        try:
            filename = f"Authentication_Records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            self.authentication_document.save(filename)
            print(f"Authentication document saved: {filename}")
            return filename
        except Exception as error:
            print(f"Error saving authentication document: {error}")
            return None

def demonstrate_advanced_calculator():
    """
    Demonstrate Question Four functionality - Advanced Calculator System
    """
    advanced_calc = AdvancedCalculator()
    
    print("\n" + "="*70)
    print("QUESTION 4: ADVANCED CALCULATOR SYSTEM DEMONSTRATION")
    print("="*70)
    
    # Record sample operations in the document
    advanced_calc.calculator_document.add_heading('Sample Mathematical Operations', 2)
    
    # Perform comprehensive sample operations
    sample_operations = [
        ('+', 15, 8),
        ('-', 25, 7),
        ('*', 6, 9),
        ('/', 45, 5),
        ('^', 2, 10),
        ('sqrt', 144, None),
        ('sin', 90, None),
        ('cos', 60, None),
        ('tan', 45, None)
    ]
    
    print("\nExecuting Sample Operations:")
    for operation, value1, value2 in sample_operations:
        if value2 is not None:
            result = advanced_calc.execute_operation(operation, value1, value2)
            print(f"{value1} {operation} {value2} = {result}")
        else:
            result = advanced_calc.execute_operation(operation, value1)
            print(f"{operation}({value1}) = {result}")
    
    # Display history in document
    advanced_calc.display_calculation_history()
    
    # Save calculator document
    calculator_filename = advanced_calc.save_calculator_document()
    
    return advanced_calc, calculator_filename

def demonstrate_student_authentication():
    """
    Demonstrate secure student authentication system
    """
    auth_system = StudentAuthenticationSystem()
    
    print("\n" + "="*70)
    print("STUDENT AUTHENTICATION SYSTEM DEMONSTRATION")
    print("="*70)
    
    # Sample student registrations
    sample_students = [
        ('john_doe', 'securePass123', '2023/U/12345', 'John Smith', '216012345'),
        ('jane_smith', 'strongPassword456', '2023/U/12346', 'Jane Doe', '216012346'),
        ('mike_wilson', 'mikePass789', '2023/U/12347', 'Mike Wilson', '216012347')
    ]
    
    # Register sample students
    print("\nStudent Registrations:")
    for username, password, reg_num, name, stud_num in sample_students:
        success, message = auth_system.register_new_student(username, password, reg_num, name, stud_num)
        print(f"Registration {username}: {message}")
    
    # Test authentication scenarios
    test_scenarios = [
        ('john_doe', 'securePass123'),
        ('jane_smith', 'wrongPassword'),
        ('nonexistent_user', 'somePassword'),
        ('mike_wilson', 'mikePass789')
    ]
    
    print("\nAuthentication Tests:")
    for username, password in test_scenarios:
        success, result = auth_system.authenticate_student(username, password)
        if success:
            print(f"✓ Login successful for {result['full_name']}")
        else:
            print(f"✗ Login failed: {result}")
    
    # Save authentication document
    auth_filename = auth_system.save_authentication_document()
    
    return auth_system, auth_filename

class StudentManagementSystem:
    """Student management system with inheritance demonstration"""
    
    def __init__(self, registration_number, full_name, student_number):
        self.registration_number = registration_number
        self.full_name = full_name
        self.student_number = student_number

    def display_student_details(self):
        """Display basic student details"""
        details = f'Student: {self.full_name}, RegNo: {self.registration_number}, StudNo: {self.student_number}'
        print(details)
        return details

class ComputerScienceStudent(StudentManagementSystem):
    """Computer Science student specialization"""
    
    course_program = 'Computer Science'
    
    def display_student_details(self):
        """Display computer science student details with inheritance"""
        base_details = super().display_student_details()
        program_info = f'Program: {ComputerScienceStudent.course_program}'
        print(program_info)
        return f"{base_details}, {program_info}"

def demonstrate_student_inheritance():
    """
    Demonstrate corrected student class program with inheritance
    """
    student_document = Document()
    student_document.add_heading('STUDENT CLASS INHERITANCE DEMONSTRATION', 0)
    student_document.add_heading('Object-Oriented Programming with Inheritance', 1)
    student_document.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    student_document.add_paragraph()
    
    student_document.add_heading('Student Objects Creation and Details', 2)
    
    # Create student objects
    regular_student = StudentManagementSystem('2023/U/12345', 'John Smith', '216012345')
    cs_student = ComputerScienceStudent('2023/U/12346', 'Jane Doe', '216012346')
    
    student_document.add_paragraph("Regular Student Details:")
    regular_details = regular_student.display_student_details()
    student_document.add_paragraph(regular_details)
    
    student_document.add_paragraph("Computer Science Student Details:")
    cs_details = cs_student.display_student_details()
    student_document.add_paragraph(cs_details)
    
    # Explanation of inheritance
    student_document.add_heading('Inheritance Explanation', 2)
    student_document.add_paragraph('''
This demonstration shows Object-Oriented Programming inheritance in action:

• StudentManagementSystem: Base class with common student attributes
• ComputerScienceStudent: Derived class that inherits from base class
• Method Overriding: display_student_details() is overridden in derived class
• super() Usage: Derived class calls base class method using super()
• Class Attributes: course_program is a class-level attribute shared by all instances
''')
    
    filename = "Student_Inheritance_Demonstration.docx"
    student_document.save(filename)
    print(f"Student inheritance document saved: {filename}")
    return student_document, filename

def generate_comprehensive_output_document():
    """
    Generate comprehensive Word document with all coursework outputs
    """
    comprehensive_document = Document()
    
    # Title Page
    comprehensive_document.add_heading('MAKERERE UNIVERSITY', 0)
    comprehensive_document.add_heading('COLLEGE OF COMPUTING AND INFORMATICS TECHNOLOGY', 1)
    comprehensive_document.add_heading('COURSEWORK 1 & 2 - COMPREHENSIVE OUTPUT', 1)
    comprehensive_document.add_paragraph()
    comprehensive_document.add_paragraph('Structured Programming and OOP / Foundations of Programming')
    comprehensive_document.add_paragraph()
    comprehensive_document.add_paragraph(f'Comprehensive Report Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    comprehensive_document.add_page_break()
    
    # Question One Summary
    comprehensive_document.add_heading('QUESTION ONE: CGPA Calculator and Basic Calculator', 1)
    comprehensive_document.add_paragraph('Implementation Features:')
    comprehensive_document.add_paragraph('• CGPA calculation using weighted averages', style='List Bullet')
    comprehensive_document.add_paragraph('• Grade classification system (Distinction, Pass, Fail)', style='List Bullet')
    comprehensive_document.add_paragraph('• Basic arithmetic operations with error handling', style='List Bullet')
    comprehensive_document.add_paragraph('• Student number extraction and processing', style='List Bullet')
    comprehensive_document.add_paragraph('• Word document integration for results storage', style='List Bullet')
    comprehensive_document.add_paragraph('• Modular code structure with separate functions', style='List Bullet')
    comprehensive_document.add_paragraph()
    
    # Question Two Summary
    comprehensive_document.add_heading('QUESTION TWO: C Structure Implementation', 1)
    comprehensive_document.add_paragraph('C Programming Concepts Demonstrated:')
    comprehensive_document.add_paragraph('• Structure declaration and initialization', style='List Bullet')
    comprehensive_document.add_paragraph('• User input handling for structure members', style='List Bullet')
    comprehensive_document.add_paragraph('• Pointer operations with structures', style='List Bullet')
    comprehensive_document.add_paragraph('• Function prototyping and parameter passing', style='List Bullet')
    comprehensive_document.add_paragraph('• Structure passing by value to functions', style='List Bullet')
    comprehensive_document.add_paragraph()
    
    # Question Three Summary
    comprehensive_document.add_heading('QUESTION THREE: Error Handling and OOP', 1)
    comprehensive_document.add_paragraph('Key Implementations:')
    comprehensive_document.add_paragraph('• Division by zero exception handling', style='List Bullet')
    comprehensive_document.add_paragraph('• File not found error handling', style='List Bullet')
    comprehensive_document.add_paragraph('• Object-Oriented Programming principles', style='List Bullet')
    comprehensive_document.add_paragraph('• Inheritance with base and derived classes', style='List Bullet')
    comprehensive_document.add_paragraph('• Encapsulation with access control', style='List Bullet')
    comprehensive_document.add_paragraph('• Polymorphism with method overriding', style='List Bullet')
    comprehensive_document.add_paragraph('• Inventory management system example', style='List Bullet')
    comprehensive_document.add_paragraph()
    
    # Question Four Summary
    comprehensive_document.add_heading('QUESTION FOUR: Advanced Calculator System', 1)
    comprehensive_document.add_paragraph('System Features:')
    comprehensive_document.add_paragraph('• Basic arithmetic operations (+, -, *, /)', style='List Bullet')
    comprehensive_document.add_paragraph('• Advanced mathematical functions (sqrt, power, trig)', style='List Bullet')
    comprehensive_document.add_paragraph('• Calculation history tracking and storage', style='List Bullet')
    comprehensive_document.add_paragraph('• Comprehensive error handling and validation', style='List Bullet')
    comprehensive_document.add_paragraph('• Object-oriented design with Calculator class', style='List Bullet')
    comprehensive_document.add_paragraph('• Secure student authentication system', style='List Bullet')
    comprehensive_document.add_paragraph('• Word document integration for all outputs', style='List Bullet')
    comprehensive_document.add_paragraph()
    
    # Programming Principles
    comprehensive_document.add_heading('Programming Principles Applied', 2)
    comprehensive_document.add_paragraph('The implementation demonstrates these software engineering principles:')
    comprehensive_document.add_paragraph('• Modular Design: Separate classes for distinct functionalities', style='List Bullet')
    comprehensive_document.add_paragraph('• Error Handling: Comprehensive try-except blocks', style='List Bullet')
    comprehensive_document.add_paragraph('• Code Reusability: Modular functions and class methods', style='List Bullet')
    comprehensive_document.add_paragraph('• Maintainability: Clear naming conventions and documentation', style='List Bullet')
    comprehensive_document.add_paragraph('• Data Encapsulation: Private methods and proper access control', style='List Bullet')
    comprehensive_document.add_paragraph('• Inheritance and Polymorphism: OOP principles in practice', style='List Bullet')
    
    filename = "Comprehensive_Coursework_Output.docx"
    comprehensive_document.save(filename)
    print(f"Comprehensive output document saved: {filename}")
    return comprehensive_document, filename

def main():
    """Main execution function for Coursework 2"""
    print("MAKERERE UNIVERSITY - COURSEWORK 2 EXECUTION")
    print("="*70)
    
    # Execute Question Four - Advanced Calculator
    calculator, calc_filename = demonstrate_advanced_calculator()
    
    # Execute Student Authentication System
    auth_system, auth_filename = demonstrate_student_authentication()
    
    # Execute Student Inheritance Demonstration
    student_doc, student_filename = demonstrate_student_inheritance()
    
    # Generate Comprehensive Output Document
    comprehensive_doc, comprehensive_filename = generate_comprehensive_output_document()
    
    print("\n" + "="*70)
    print("COURSEWORK 2 EXECUTION COMPLETED SUCCESSFULLY!")
    print("="*70)
    print(f"Advanced Calculator: {calc_filename}")
    print(f"Authentication System: {auth_filename}")
    print(f"Student Inheritance: {student_filename}")
    print(f"Comprehensive Output: {comprehensive_filename}")
    print("="*70)

if __name__ == "__main__":
    main()