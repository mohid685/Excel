#include <iostream>
#include <vector>
#include <conio.h>
#include <Windows.h>
#include <string>
#include <iomanip>
#include <algorithm>
#include <fstream>
#include <sstream>
#include <ostream>
#include "Excel.h"

#define BROWN 6
#define GREEN 2
#define RED 4
#define LBLUE 9
#define YELLOW 14
#define GRAY 8
#define PINK 13
#define ORANGE 12
#define GREEN 10
#define BLUE 9
#define PURPLE 5

using namespace std;
char sym = -37;

struct CellData {
    int intData;
    float floatData;
     string strData;
};
enum DataType {
    INT_TYPE,
    FLOAT_TYPE,
    STRING_TYPE,
    NONE
};
struct CellInfo {
    DataType dataType;
    CellData cellData;
};
class Cell {
public:
    CellData data;
    DataType type;
    Cell* up;
    Cell* down;
    Cell* left;
    Cell* right;

    Cell() : type(NONE), up(NULL), down(NULL), left(NULL), right(NULL) {}
};
class Excel {
private:
    Cell* root;      
    Cell* current;
    vector<int> rowCoords;
    vector<int> colCoords;
    vector<CellData> copiedData;
public:
    Excel() {
        setConsoleColor(GREEN);
        root = new Cell(); 
        current = root;

        Cell* rowStart = root;  
        Cell* prevRow = nullptr;  

        for (int i = 0; i < 5; i++) {
            Cell* currentRow = rowStart;
            Cell* prevRowCell = prevRow; 

            for (int j = 0; j < 5; j++) {
                if (prevRowCell != nullptr) {
                    currentRow->up = prevRowCell;
                    prevRowCell->down = currentRow;
                    prevRowCell = prevRowCell->right; 
                }

                if (j < 4) {
                    currentRow->right = new Cell();
                    currentRow->right->left = currentRow;
                    currentRow = currentRow->right;
                }
            }

            if (i < 4) {
                rowStart->down = new Cell();
                rowStart->down->up = rowStart;
                prevRow = rowStart; 
                rowStart = rowStart->down;
            }
        }
    }
    void setConsoleColor(int color) {
        SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), color);
    }
    void moveCurrent(char direction) {
        Cell* previous = current; 

        switch (direction) {
        case 'U':
            if (current->up) current = current->up;
            else {
                cout << "Reached the upper boundary! Stay within the grid." <<"\n";
            }
            break;
        case 'D':
            if (current->down) current = current->down;
            else {
                cout << "Reached the lower boundary! Stay within the grid." <<"\n";
            }
            break;
        case 'L':
            if (current->left) current = current->left;
            else {
                cout << "Reached the left boundary! Stay within the grid." <<"\n";
            }
            break;
        case 'R':
            if (current->right) current = current->right;
            else {
                cout << "Reached the right boundary! Stay within the grid." <<"\n";
            }
            break;
        default:
            cout << "Invalid direction!" <<"\n";
        }

        if (current == previous) {
            cout << "Returning to previous cell." <<"\n";
        }
    }
    void gotoRowCol(int rpos, int cpos)
    {
        COORD scrn;
        HANDLE hOuput = GetStdHandle(STD_OUTPUT_HANDLE);
        scrn.X = cpos;
        scrn.Y = rpos;
        SetConsoleCursorPosition(hOuput, scrn);
    }
    void displayGrid() {
        Cell* row = root;
        char ch = -37; 

        Cell* temp = row;
        while (temp) {
            setConsoleColor(7); 
            cout << "ßßßßß";
            temp = temp->right;
        }
        cout << "\n";

        while (row) {
            Cell* col = row;

            if (col == current) {
                setConsoleColor(GREEN); 
            }
            else {
                setConsoleColor(7); 
            }
            cout << ch; 

            while (col) {
                if (col == current) {
                    setConsoleColor(GREEN); 
                }

                switch (col->type) {
                case INT_TYPE:
                    cout << setw(4) << col->data.intData;
                    break;
                case FLOAT_TYPE:
                    cout << setw(4) << col->data.floatData;
                    break;
                case STRING_TYPE:
                    cout << setw(4) << col->data.strData;
                    break;
                case NONE:
                    cout << "    "; 
                    break;
                }

                if (col->right && col->right == current) {
                    setConsoleColor(GREEN); 
                }
                else {
                    setConsoleColor(7); 
                }

                cout << ch;

                col = col->right;
            }

            cout << endl;

            col = row;
            while (col) {
                setConsoleColor(7); 
                cout << "ßßßßß";
                col = col->right;
            }
            cout << endl;

            row = row->down;
        }

        setConsoleColor(7); 
    }
    void setCurrentData() {
        setConsoleColor(GREEN);
        char choice;
         cout << "Enter the type of data you want to insert (i = int, f = float, s = string): ";
         cin >> choice;

        int intVal = 0;
        float floatVal = 0.0f;
         string strVal;

        switch (choice) {
        case 'i':
             cout << "Enter an integer value: ";
             cin >> intVal;
            current->data.intData = intVal;
            current->type = INT_TYPE;
            break;

        case 'f':
             cout << "Enter a float value: ";
             cin >> floatVal;
            current->data.floatData = floatVal;
            current->type = FLOAT_TYPE;
            break;

        case 's':
             cout << "Enter a string (max 4 characters): ";
             cin >> strVal;
            if (strVal.length() > 4) {
                 cout << "String too long! Enter a string with max 4 characters." <<  endl;
            }
            else {
                current->data.strData = strVal;
                current->type = STRING_TYPE;
            }
            break;

        default:
             cout << "Invalid data type!" <<  endl;
        }
        setConsoleColor(7);
    }
    void addRowAboveMost() {
        Cell* currentRow = current;
        while (currentRow->left) {
            currentRow = currentRow->left;
        }

        while (currentRow->up) {
            currentRow = currentRow->up;
        }

        Cell* newRowHead = new Cell();
        Cell* newRowCurrent = newRowHead;
        Cell* temp = currentRow;

        while (temp) {
            Cell* newCell = new Cell();
            newCell->down = temp;
            temp->up = newCell;

            if (temp->left) {
                newCell->left = temp->left->up;
                newCell->left->right = newCell;
            }

            temp = temp->right;
            if (temp) {
                newRowCurrent->right = new Cell();
                newRowCurrent->right->left = newRowCurrent;
                newRowCurrent = newRowCurrent->right;
            }
        }

        newRowHead->down = currentRow;
        if (currentRow) {
            currentRow->up = newRowHead;
        }

        Cell* newRowIterator = newRowHead->right;
        Cell* rowIterator = currentRow->right;

        while (newRowIterator && rowIterator) {
            newRowIterator->down = rowIterator;
            rowIterator->up = newRowIterator;

            newRowIterator = newRowIterator->right;
            rowIterator = rowIterator->right;
        }

        if (currentRow == root) {
            root = newRowHead;
        }
    }
    void addRowAbove() {
        Cell* currentRow = current;

        while (currentRow->left) {
            currentRow = currentRow->left;
        }

        Cell* aboveRow = currentRow->up;

        Cell* newRowHead = new Cell();
        Cell* newRowCurrent = newRowHead;
        Cell* temp = currentRow;

        while (temp) {
            Cell* newCell = new Cell();
            newCell->down = temp;
            temp->up = newCell;

            if (temp->left) {
                newCell->left = temp->left->up;
                if (newCell->left) {
                    newCell->left->right = newCell;
                }
            }

            temp = temp->right;

            if (temp) {
                newRowCurrent->right = new Cell();
                newRowCurrent->right->left = newRowCurrent;
                newRowCurrent = newRowCurrent->right;
            }
        }

        newRowHead->down = currentRow;
        if (currentRow) {
            currentRow->up = newRowHead;
        }

        if (aboveRow) {
            newRowHead->up = aboveRow;
            Cell* aboveRowIterator = aboveRow;
            Cell* newRowIterator = newRowHead;

            while (aboveRowIterator && newRowIterator) {
                newRowIterator->up = aboveRowIterator;
                aboveRowIterator->down = newRowIterator;

                aboveRowIterator = aboveRowIterator->right;
                newRowIterator = newRowIterator->right;
            }
        }

        if (!aboveRow) {
            root = newRowHead;
        }

        //current = newRowHead;

        Cell* newRowIterator = newRowHead->right;
        Cell* rowIterator = currentRow->right;

        while (newRowIterator && rowIterator) {
            newRowIterator->down = rowIterator;
            rowIterator->up = newRowIterator;

            newRowIterator = newRowIterator->right;
            rowIterator = rowIterator->right;
        }
    }
    void addRowBelowMost() {
        Cell* currentRow = current;

        while (currentRow->left) {
            currentRow = currentRow->left;
        }
        while (currentRow->down) {
            currentRow = currentRow->down;
        }

        Cell* newRow = new Cell();
        Cell* temp = currentRow;

        while (temp) {
            Cell* newCell = new Cell();
            newCell->up = temp;
            temp->down = newCell;

            if (temp->left) {
                newCell->left = temp->left->down;
                temp->left->down->right = newCell;
            }

            temp = temp->right;
        }
    }
    void addRowBelow() {
        Cell* currentRow = current;

        while (currentRow->left) {
            currentRow = currentRow->left;
        }

        Cell* belowRow = currentRow->down;

        Cell* newRowHead = new Cell();
        Cell* newRowCurrent = newRowHead;
        Cell* temp = currentRow;

        while (temp) {
            Cell* newCell = new Cell();
            newCell->up = temp;
            temp->down = newCell;

            if (temp->left) {
                newCell->left = temp->left->down;
                if (newCell->left) {
                    newCell->left->right = newCell;
                }
            }

            temp = temp->right;

            if (temp) {
                newRowCurrent->right = new Cell();
                newRowCurrent->right->left = newRowCurrent;
                newRowCurrent = newRowCurrent->right;
            }
        }

        newRowHead->up = currentRow;
        if (currentRow) {
            currentRow->down = newRowHead;
        }

        if (belowRow) {
            newRowHead->down = belowRow;
            Cell* belowRowIterator = belowRow;
            Cell* newRowIterator = newRowHead;

            while (belowRowIterator && newRowIterator) {
                newRowIterator->down = belowRowIterator;
                belowRowIterator->up = newRowIterator;

                belowRowIterator = belowRowIterator->right;
                newRowIterator = newRowIterator->right;
            }
        }

        //current = newRowHead;

        Cell* newRowIterator = newRowHead->right;
        Cell* rowIterator = currentRow->right;

        while (newRowIterator && rowIterator) {
            newRowIterator->up = rowIterator;
            rowIterator->down = newRowIterator;

            newRowIterator = newRowIterator->right;
            rowIterator = rowIterator->right;
        }
    }
    void addColLeftMost() {
        Cell* currentCol = current;
        while (currentCol->left) {
            currentCol = currentCol->left;
        }

        while (currentCol->up) {
            currentCol = currentCol->up;
        }

        Cell* newColHead = new Cell();
        Cell* newColCurrent = newColHead;
        Cell* temp = currentCol;

        while (temp) {
            Cell* newCell = new Cell();
            newCell->right = temp;
            temp->left = newCell;

            if (temp->up) {
                newCell->up = temp->up->left;
                newCell->up->down = newCell;
            }

            temp = temp->down;
            if (temp) {
                newColCurrent->down = new Cell();
                newColCurrent->down->up = newColCurrent;
                newColCurrent = newColCurrent->down;
            }
        }

        newColHead->right = currentCol;
        if (currentCol) {
            currentCol->left = newColHead;
        }

        Cell* newColIterator = newColHead->down;
        Cell* colIterator = currentCol->down;

        while (newColIterator && colIterator) {
            newColIterator->right = colIterator;
            colIterator->left = newColIterator;

            newColIterator = newColIterator->down;
            colIterator = colIterator->down;
        }

        if (currentCol == root) {
            root = newColHead;
        }
    }
    void addColLeft() {
        Cell* currentCol = current;

        if (current->left == NULL) {
            addColLeftMost();
            return;
        }

        while (currentCol->up) {
            currentCol = currentCol->up;
        }

        Cell* leftCol = currentCol->left;

        Cell* newColHead = new Cell();
        Cell* newColCurrent = newColHead;
        Cell* temp = currentCol;

        while (temp) {
            Cell* newCell = new Cell();
            newCell->right = temp;   
            temp->left = newCell;    

            if (newColCurrent->up) {
                newColCurrent->up->down = newColCurrent;
            }

            temp = temp->down;

            if (temp) {
                newColCurrent->down = new Cell();
                newColCurrent->down->up = newColCurrent;
                newColCurrent = newColCurrent->down;
            }
        }

        newColHead->right = currentCol;
        if (currentCol) {
            currentCol->left = newColHead;
        }

        if (leftCol) {
            Cell* newColIterator = newColHead;
            Cell* leftColIterator = leftCol;
            Cell* currentColIterator = currentCol;

            while (newColIterator) {
                newColIterator->left = leftColIterator;
                if (leftColIterator) {
                    leftColIterator->right = newColIterator;
                    leftColIterator = leftColIterator->down;
                }

                if (currentColIterator) {
                    currentColIterator->left = newColIterator;
                    newColIterator->right = currentColIterator;
                    currentColIterator = currentColIterator->down;
                }

                newColIterator = newColIterator->down;
            }
        }
        //current = newColHead;
    }
    void addColRightMost() {
        Cell* currentCol = current;
        while (currentCol->right) {
            currentCol = currentCol->right;
        }

        while (currentCol->up) {
            currentCol = currentCol->up;
        }

        Cell* newColHead = new Cell();
        Cell* newColCurrent = newColHead;
        Cell* temp = currentCol;

        while (temp) {
            Cell* newCell = new Cell();
            newCell->left = temp;
            temp->right = newCell;

            if (temp->up) {
                newCell->up = temp->up->right;
                newCell->up->down = newCell;
            }

            temp = temp->down;
            if (temp) {
                newColCurrent->down = new Cell();
                newColCurrent->down->up = newColCurrent;
                newColCurrent = newColCurrent->down;
            }
        }

        newColHead->left = currentCol;
        if (currentCol) {
            currentCol->right = newColHead;
        }

        Cell* newColIterator = newColHead->down;
        Cell* colIterator = currentCol->down;

        while (newColIterator && colIterator) {
            newColIterator->left = colIterator;
            colIterator->right = newColIterator;

            newColIterator = newColIterator->down;
            colIterator = colIterator->down;
        }
    }
    void addColRight() {
        Cell* currentCol = current;
        while (currentCol->up) {
            currentCol = currentCol->up;
        }

        Cell* rightCol = currentCol->right;

        Cell* newColHead = new Cell();
        Cell* newColCurrent = newColHead;
        Cell* temp = currentCol;

        while (temp) {
            Cell* newCell = new Cell();
            newCell->left = temp;   
            temp->right = newCell;  

            if (newColCurrent->up) {
                newColCurrent->up->down = newColCurrent;
            }

            temp = temp->down;

            if (temp) {
                newColCurrent->down = new Cell();
                newColCurrent->down->up = newColCurrent;
                newColCurrent = newColCurrent->down;
            }
        }

        newColHead->left = currentCol;
        if (currentCol) {
            currentCol->right = newColHead;
        }

        if (rightCol) {
            Cell* newColIterator = newColHead;
            Cell* rightColIterator = rightCol;
            Cell* currentColIterator = currentCol;

            while (newColIterator) {
                newColIterator->right = rightColIterator;
                if (rightColIterator) {
                    rightColIterator->left = newColIterator;
                    rightColIterator = rightColIterator->down;
                }

                if (currentColIterator) {
                    currentColIterator->right = newColIterator;
                    newColIterator->left = currentColIterator;
                    currentColIterator = currentColIterator->down;
                }

                newColIterator = newColIterator->down;
            }
        }
        //current = newColHead;
    }
    void InsertCellByRightShift() {
        Cell* temp = current;   
        Cell* rightmost = temp;

        while (rightmost->right) {
            rightmost = rightmost->right;
        }

        addColRightMost();
        rightmost = rightmost->right;  

        while (rightmost != current) {
            rightmost->data = rightmost->left->data;  
            rightmost->type = rightmost->left->type;  
            rightmost = rightmost->left;
        }

        current->data = {};   
        current->type = NONE; 
    }
    void InsertCellByDownShift() {
        Cell* temp = current;   
        Cell* bottommost = temp;

        while (bottommost->down) {
            bottommost = bottommost->down;
        }

        addRowBelowMost();
        bottommost = bottommost->down; 

        while (bottommost != current) {
            bottommost->data = bottommost->up->data;
            bottommost->type = bottommost->up->type;  
            bottommost = bottommost->up;
        }

        current->data = {};   
        current->type = NONE; 
    }
    void DeleteCellByLeftShift() {
        Cell* temp = current;  
        Cell* rightmost = temp;
        Cell* tracker = current;

        while (rightmost->right) {
            rightmost = rightmost->right;
        }

        while (temp != rightmost) {
            temp->data = temp->right->data;  
            temp->type = temp->right->type;  
            temp = temp->right;
        }

        rightmost->data = {};   
        rightmost->type = NONE; 

        current = current->left;

        if (current == NULL) {
            current = tracker;
        }
    }
    void deleteCellByUpShift() {

        Cell* temp = current;
        Cell* bottomMost = temp;
        Cell* tracker = current;

        while (bottomMost->down) {
            bottomMost = bottomMost->down;
        }

        while (temp != bottomMost) {
            temp->data = temp->down->data;
            temp->type = temp->down->type;
            temp = temp->down;
        }

        bottomMost->data = {};
        bottomMost->type = NONE;

        current = current->up;

        if (current == NULL) {
            current = tracker;
        }
    }
    void deleteRow() {

        if (current == NULL) {
            throw runtime_error("No Rows To Delete!!!\n");
            return;
        }

        Cell* currentRow = current;
        while (currentRow->left) {
            currentRow = currentRow->left;
        }

        if (current->up == NULL) {
            current = current->down;
        }
        else if (current->down == NULL) {
            current = current->up;
        }
        else if (current->up != NULL && current->down != NULL) {
            current = current->up;
        }

        while (currentRow) {

            Cell* temp = currentRow->right;

            if (currentRow->up) {
                currentRow->up->down = currentRow->down;
            }
            if (currentRow->down) {
                currentRow->down->up = currentRow->up;
            }

            if (currentRow == root) {
                root = root->down;
            }

            delete currentRow;
            currentRow = temp;
        }
    }
    void deleteColumn() {
        if (current == NULL) {
            throw runtime_error("No Columns To Delete!!!\n");
            return;
        }
        Cell* tempRoot = root;
        Cell* currentCol = current;
        while (currentCol->up) {
            currentCol = currentCol->up;
        }

        if (current->left == NULL) {
            current = current->right;  
        }
        else if (current->right == NULL) {
            current = current->left; 
        }
        else if (current->left != NULL && current->right != NULL) {
            current = current->right;  
        }

        while (currentCol) {
            Cell* temp = currentCol->down;

            if (currentCol->left) {
                currentCol->left->right = currentCol->right;
            }
            if (currentCol->right) {
                currentCol->right->left = currentCol->left;
            }

            if (currentCol == root) {
                root = root->right;
            }

            delete currentCol;
            currentCol = temp;
        }
    }
    void ClearRow() {
        Cell* currentRow = current;
        while (currentRow->left) {
            currentRow = currentRow->left;
        }

        while (currentRow) {
            currentRow->data.intData = 0;
            currentRow->data.floatData = 0.0f;
            currentRow->data.strData = "";
            currentRow->type = NONE;  

            currentRow = currentRow->right;
        }
    }
    void ClearCols() {

        Cell* currentCol = current;
        while (currentCol->up) {
            currentCol = currentCol->up;
        }

        while (currentCol) {

            currentCol->data.intData = 0;
            currentCol->data.floatData = 0.0f;
            currentCol->data.strData = "";
            currentCol->type = NONE; 

            currentCol = currentCol->down;
        }
    }
    Cell* getCellByCoordinates(int x, int y) {
        if (x < 0 || y < 0) {
             cout << "Error: Invalid coordinates!" <<  endl;
            return nullptr;
        }

        Cell* temp = root;  
      
        for (int i = 0; i < y && temp != nullptr; i++) {
            temp = temp->down;
        }

        for (int j = 0; j < x && temp != nullptr; j++) {
            temp = temp->right;
        }

        if (temp == nullptr) {
             cout << "Error: Coordinates out of bounds!" <<  endl;
        }
        return temp;
    }
    float RangeSum(Cell* start, Cell* end) {
        if (!start || !end) {
            cout << "Error: Invalid range or grid is empty!" << endl;
            return -1;
        }

        int startRow = 0, startCol = 0, endRow = 0, endCol = 0;

        Cell* temp = start;
        while (temp->up) {
            temp = temp->up;
            startRow++;
        }
        while (temp->left) {
            temp = temp->left;
            startCol++;
        }
        temp = end;
        while (temp->up) {
            temp = temp->up;
            endRow++;
        }
        while (temp->left) {
            temp = temp->left;
            endCol++;
        }

        Cell* topLeft = start;
        while (topLeft->up) { 
            topLeft = topLeft->up; 
        }
        while (topLeft->left) {
            topLeft = topLeft->left;
        }

        Cell* bottomRight = end;
        while (bottomRight->down) {
            bottomRight = bottomRight->down;
        } 
        while (bottomRight->right) {
            bottomRight = bottomRight->right;
        } 

        if (startRow > endRow || startCol > endCol) {
            cout << "Error: Invalid range!" << endl;
            return -1;
        }
        float sum = 0;
        for (Cell* rowPtr = topLeft; rowPtr && startRow <= endRow; rowPtr = rowPtr->down, startRow++) {
            int currentCol = startCol;
            for (Cell* colPtr = rowPtr; colPtr && currentCol <= endCol; colPtr = colPtr->right, currentCol++) {
                if (colPtr->type == INT_TYPE) {
                    sum += colPtr->data.intData;
                }
                else if (colPtr->type == FLOAT_TYPE) {
                    sum += colPtr->data.floatData;
                }
            }
        }
        return sum;
    }
    float RangeAverage(Cell* start, Cell* end) {
        if (!start || !end) {
            cout << "Error: Invalid range or grid is empty!" << endl;
            return -1;
        }

        int startRow = 0, startCol = 0, endRow = 0, endCol = 0;

        Cell* temp = start;
        while (temp->up) {
            temp = temp->up;
            startRow++;
        }
        while (temp->left) {
            temp = temp->left;
            startCol++;
        }
        temp = end;
        while (temp->up) {
            temp = temp->up;
            endRow++;
        }
        while (temp->left) {
            temp = temp->left;
            endCol++;
        }

        Cell* topLeft = start;
        while (topLeft->up) {
            topLeft = topLeft->up;
        }
        while (topLeft->left) {
            topLeft = topLeft->left;
        }

        Cell* bottomRight = end;
        while (bottomRight->down) {
            bottomRight = bottomRight->down;
        }
        while (bottomRight->right) {
            bottomRight = bottomRight->right;
        }

        if (startRow > endRow || startCol > endCol) {
            cout << "Error: Invalid range!" << endl;
            return -1;
        }
        float sum = 0;
        int count = 0;
        for (Cell* rowPtr = topLeft; rowPtr && startRow <= endRow; rowPtr = rowPtr->down, startRow++) {
            int currentCol = startCol;
            for (Cell* colPtr = rowPtr; colPtr && currentCol <= endCol; colPtr = colPtr->right, currentCol++) {
                if (colPtr->type == INT_TYPE) {
                    sum += colPtr->data.intData;
                    count++;
                }
                else if (colPtr->type == FLOAT_TYPE) {
                    sum += colPtr->data.floatData;
                    count++;
                }
                
            }
        }
        return sum/count;
    }
    int CountCellsInRange(Cell* start, Cell* end) {

        if (!start || !end) {
            cout << "Error: Invalid start or end cell!" << endl;
            return -1;
        }

        Cell* topLeft = start;
        Cell* bottomRight = end;

        while (topLeft->left && (topLeft != bottomRight)) {
            topLeft = topLeft->left;
        } 
        while (topLeft->up && (topLeft != bottomRight)) {
            topLeft = topLeft->up;
        } 

        while (bottomRight->right && (bottomRight != topLeft)) {
            bottomRight = bottomRight->right;
        }
        while (bottomRight->down && (bottomRight != topLeft)) { 
            bottomRight = bottomRight->down; 
        }

        int count = 0;
        Cell* currentRow = topLeft;
        while (currentRow && currentRow != bottomRight->down) {
            Cell* currentCell = currentRow;
            while (currentCell && currentCell != bottomRight->right) {
                if (currentCell->type == INT_TYPE || currentCell->type == FLOAT_TYPE || currentCell->type == STRING_TYPE) {
                    count++;
                }
                currentCell = currentCell->right;
            }
            currentRow = currentRow->down;
        }

        return count;
    }
    vector<float> rangeMaxMin(Cell* start, Cell* end) {

        if (!start || !end) {
            cout << "Error: Invalid range or grid is empty!" << endl;
            return {};
        }

        int startRow = 0, startCol = 0, endRow = 0, endCol = 0;

        Cell* temp = start;
        while (temp->up) {
            temp = temp->up;
            startRow++;
        }
        while (temp->left) {
            temp = temp->left;
            startCol++;
        }
        temp = end;
        while (temp->up) {
            temp = temp->up;
            endRow++;
        }
        while (temp->left) {
            temp = temp->left;
            endCol++;
        }

        Cell* topLeft = start;
        while (topLeft->up) {
            topLeft = topLeft->up;
        }
        while (topLeft->left) {
            topLeft = topLeft->left;
        }

        Cell* bottomRight = end;
        while (bottomRight->down) {
            bottomRight = bottomRight->down;
        }
        while (bottomRight->right) {
            bottomRight = bottomRight->right;
        }

        if (startRow > endRow || startCol > endCol) {
            cout << "Error: Invalid range!" << endl;
            return {};
        }
        
        float maxi = FLT_MIN;
        float mini = FLT_MAX;

        for (Cell* rowPtr = topLeft; rowPtr && startRow <= endRow; rowPtr = rowPtr->down, startRow++) {
            int currentCol = startCol;
            for (Cell* colPtr = rowPtr; colPtr && currentCol <= endCol; colPtr = colPtr->right, currentCol++) {
                if (colPtr->type == INT_TYPE) {

                    if (colPtr->data.intData > maxi) {
                        maxi = colPtr->data.intData;
                    }
                    if (colPtr->data.intData < mini) {
                        mini = colPtr->data.intData;
                    }
                }
                else if (colPtr->type == FLOAT_TYPE) {
                    if (colPtr->data.floatData > maxi) {
                        maxi = colPtr->data.floatData;
                    }
                    if (colPtr->data.floatData < mini) {
                        mini = colPtr->data.floatData;
                    }
                }
            }
        }

        vector<float> result;
        result.push_back(maxi);
        result.push_back(mini);

        return result;

    }
    void CopyCells(Cell* start, Cell* end) {
        if (!start || !end) {
            cout << "Error: The grid or selected range is empty!" << endl;
            return;
        }
        rowCoords.clear();
        colCoords.clear();
        copiedData.clear();

        Cell* topLeft = start;
        Cell* bottomRight = end;

        while (topLeft->left && (topLeft != bottomRight)) {
            topLeft = topLeft->left;
        }
        while (topLeft->up && (topLeft != bottomRight)) {
            topLeft = topLeft->up;
        }

        while (bottomRight->right && (bottomRight != topLeft)) {
            bottomRight = bottomRight->right;
        }
        while (bottomRight->down && (bottomRight != topLeft)) {
            bottomRight = bottomRight->down;
        }

        int row = 0, col = 0;
        Cell* currentRow = topLeft;

        while (currentRow) {
            Cell* currentCell = currentRow;
            col = 0;  

            while (currentCell) {
                copiedData.push_back(currentCell->data);
                rowCoords.push_back(row);
                colCoords.push_back(col);

                if (currentCell == end) {
                    cout << "Cells copied successfully!" << "\n";
                    return;
                }

                currentCell = currentCell->right;
                col++;
            }

            currentRow = currentRow->down;
            row++;
        }
        cout << "Cells copied successfully!" << "\n";
    }
    void CutCells(Cell* start, Cell* end) {
        if (!start || !end) {
            cout << "Error: The grid or selected range is empty!" << endl;
            return;
        }
        CopyCells(start, end);
        Cell* topLeft = start;
        Cell* bottomRight = end;

        int row = 0, col = 0;
        Cell* currentRow = topLeft;
        while (currentRow) {
            Cell* currentCell = currentRow;
            col = 0;  

            while (currentCell) {
                currentCell->data.intData = 0;
                currentCell->data.floatData = 0.0f;
                currentCell->data.strData = "";
                currentCell->type = NONE; 
                if (currentCell == end) {
                    cout << "Cells cut successfully!" << "\n";
                    return;
                }
                currentCell = currentCell->right;
                col++;
            }
            currentRow = currentRow->down;
            row++;
        }
        cout << "Cells cut successfully!" << "\n";
    }
    DataType GetDataType(CellData data) {
        if (data.intData != 0) {
            return INT_TYPE;
        }
        else if (data.floatData != 0.0f) {
            return FLOAT_TYPE;
        }
        else if (!data.strData.empty()) {
            return STRING_TYPE;
        }
        return NONE;
    }
    void PasteCells() {
        if (!current) {
            cout << "Error: No valid current cell for pasting!" << endl;
            return;
        }

        if (copiedData.empty()) {
            cout << "No data to paste!" << "\n";
            return;
        }

        int maxRow = rowCoords[0];
        int maxCol = colCoords[0];

        for (int i = 1; i < rowCoords.size(); i++) {
            if (rowCoords[i] > maxRow) {
                maxRow = rowCoords[i];
            }
        }

        for (int i = 1; i < colCoords.size(); i++) {
            if (colCoords[i] > maxCol) {
                maxCol = colCoords[i];
            }
        }


        Cell* temp = current;
        int rowsAvailable = 0, colsAvailable = 0;

        Cell* tempRow = current;
        while (tempRow) {
            rowsAvailable++;
            tempRow = tempRow->down;
        }

        tempRow = current;
        while (tempRow) {
            Cell* tempCol = tempRow;
            colsAvailable = 0;
            while (tempCol) {
                colsAvailable++;
                tempCol = tempCol->right;
            }
            tempRow = tempRow->down;
        }

        while (rowsAvailable <= maxRow) {
            addRowBelowMost();
            rowsAvailable++;
        }

        while (colsAvailable <= maxCol) {
            addColRightMost();
            colsAvailable++;
        }

        tempRow = current;
        for (int i = 0; i < copiedData.size(); i++) {
            int targetRow = rowCoords[i];
            int targetCol = colCoords[i];

            Cell* targetCell = current;
            for (int r = 0; r < targetRow; r++) {
                if (targetCell->down) {
                    targetCell = targetCell->down;
                }
            }
            for (int c = 0; c < targetCol; c++) {
                if (targetCell->right) {
                    targetCell = targetCell->right;
                }
            }

            targetCell->data = copiedData[i];
            targetCell->type = GetDataType(copiedData[i]);
        }

        cout << "Cells pasted successfully!" << "\n";
    }
    void testCopyCut() {
        int startX, startY, endX, endY;
        char choice;

        cout << "Enter the starting cell coordinates (startX, startY): ";
        cin >> startX >> startY;

        cout << "Enter the ending cell coordinates (endX, endY): ";
        cin >> endX >> endY;

        Cell* start = getCellByCoordinates(startX, startY);
        Cell* end = getCellByCoordinates(endX, endY);

        cout << "Do you want to (c)opy or (x)cut? ";
        cin >> choice;

        if (choice == 'c' || choice == 'C') {
            CopyCells(start, end);
            cout << "Copied data:" <<"\n";
            for (int i = 0; i < copiedData.size(); i++) {
                cout << "Cell at (" << rowCoords[i] << ", " << colCoords[i] << ") contains: ";
                if (copiedData[i].intData != 0)  
                    cout << copiedData[i].intData <<"\n";
                else if (copiedData[i].floatData != 0.0f)  
                    cout << copiedData[i].floatData <<"\n";
                else if (!copiedData[i].strData.empty())  
                    cout << copiedData[i].strData <<"\n";
            }
        }
        else if (choice == 'x' || choice == 'X') {
            CutCells(start, end);
            cout << "Cut data:" <<"\n";
            for (int i = 0; i < copiedData.size(); i++) {
                cout << "Cell at (" << rowCoords[i] << ", " << colCoords[i] << ") contains: ";
                if (copiedData[i].intData != 0)
                    cout << copiedData[i].intData <<"\n";
                else if (copiedData[i].floatData != 0.0f)
                    cout << copiedData[i].floatData <<"\n";
                else if (!copiedData[i].strData.empty())
                    cout << copiedData[i].strData <<"\n";
            }
        }

        _getch();  
    }
    void testPaste() {
        PasteCells();
        cout << "Data pasted successfully." <<"\n";
        _getch();
    }
    void testCountCells() {
        int startX, startY, endX, endY;
         cout << "Enter the starting cell coordinates (startX, startY): ";
         cin >> startX >> startY;

         cout << "Enter the ending cell coordinates (endX, endY): ";
         cin >> endX >> endY;

   
        Cell* start = getCellByCoordinates(startX, startY);
        Cell* end = getCellByCoordinates(endX, endY);

       
        int count = CountCellsInRange(start, end);
        if (count != -1) {
             cout << "Number of cells with data in the selected range: " << count <<  endl;
        }
    }
    void testMaxMin() {

        int startX, startY, endX, endY;
         cout << "Enter the starting cell coordinates (startX, startY): ";
         cin >> startX >> startY;

         cout << "Enter the ending cell coordinates (endX, endY): ";
         cin >> endX >> endY;

        Cell* start = getCellByCoordinates(startX, startY);
        Cell* end = getCellByCoordinates(endX, endY);

        vector<float> out = rangeMaxMin(start, end);

         cout << "Minimum value in the range: " << out[1] <<  endl;
         cout << "Maximum value in the range: " << out[0] <<  endl;

    }
    void testRangeSum() {
        Cell* start = nullptr;
        Cell* end = nullptr;
        char choice;
        cout << "x represents cols 0 based indexed\n";
        cout << "y represents rows 0 based indexed\n";

        int startX, startY, endX, endY;
        cout << "Enter starting cell (x, y): ";

        cin >> startX >> startY;

        cout << "Enter ending cell (x, y): ";
        cin >> endX >> endY;

        start = getCellByCoordinates(startX, startY);
        end = getCellByCoordinates(endX, endY);


        if (start != nullptr && end != nullptr) {
            float sum = RangeSum(start, end);
            cout << "The range sum is: " << sum <<"\n";
        }
        else {
            cout << "Error: Invalid start or end cells!" <<"\n";
        }
    }
    void testRangeAvg() {
        Cell* start = nullptr;
        Cell* end = nullptr;
      
        cout << "x represents cols 0 based indexed\n";
        cout << "y represents rows 0 based indexed\n";

        int startX, startY, endX, endY;
        cout << "Enter starting cell (x, y): ";

        cin >> startX >> startY;

        cout << "Enter ending cell (x, y): ";
        cin >> endX >> endY;

        start = getCellByCoordinates(startX, startY);
        end = getCellByCoordinates(endX, endY);


        if (start != nullptr && end != nullptr) {
            float avg = RangeAverage(start, end);
            cout << "The range average is: " << avg <<"\n";
        }
        else {
            cout << "Error: Invalid start or end cells!" <<"\n";
        }
    }
    void boxDisplay(int cr, int cc, int br, int bc) {
        for (int r = 0; r < br; r++) {
            for (int c = 0; c < bc; c++) {
                if (r == 0 || r == br - 1 || c == 0 || c == bc - 1) {

                    gotoRowCol(cr + r, cc + c);
                    cout << sym;
                }
            }
        }
    }
    void gridDisplay(int width, int height, int rows, int cols) {

        int brs = height / rows;
        int bcs = width / cols;

        for (int r = 0; r < rows; r++) {

            for (int c = 0; c < cols; c++) {

                boxDisplay(r * brs, c * bcs, brs, bcs);
            }
        }
    }
    void displayCommandShortcuts() {
        system("CLS");
        setConsoleColor(GREEN);
        vector<string> commands = {
            "T: Add Row Up",
            "V: Add Row Down",
            "G: Add Column Right",
            "F: Add Column Left",
            "A: Delete Row",
            "Z: Delete Column",
            "O: Insert Right Shift",
            "P: Insert Down Shift",
            "K: Delete Left Shift",
            "L: Delete Up Shift",
            "R: Range Sum",
            "U: Range Avg",
            "W: Range Count",
            "Q: Range Max-Min",
            "B: Copy-Cut",
            "N: Paste",
            "1: Clear Row",
            "2: Clear Column",
            "5: Save Sheet",
            "6: Load Sheet"
        };

        int screenWidth = 80;
        int boxWidth = 50;
        int boxHeight = commands.size() + 4;
        int startRow = (25 - boxHeight) / 2;
        int startCol = (screenWidth - boxWidth) / 2;

        gotoRowCol(startRow, startCol);
        cout << "-" << string(boxWidth - 2, '-') << "-";

        for (int i = 1; i < boxHeight - 1; i++) {
            gotoRowCol(startRow + i, startCol);
            cout << "|" << string(boxWidth - 2, ' ') << "|";
        }

        gotoRowCol(startRow + boxHeight - 1, startCol);
        cout << "-" << string(boxWidth - 2, '-') << "-";

        for (int i = 0; i < commands.size(); i++) {
            gotoRowCol(startRow + i + 2, startCol + 2);  
            cout << commands[i];
        }

        gotoRowCol(startRow + boxHeight, startCol + (boxWidth / 2) - 10);
        cout << "Press any key to continue...";
        _getch();
        setConsoleColor(7);
        system("CLS");
        displayGrid();
    }
    void saveSheet(const string fileName) {

        ofstream wrt(fileName);

        if (!wrt) {
            cout << "Cannot Save The File!!\n\n";
            return;
        }

        Cell* rowIter = root;

        while (rowIter != NULL) {

            Cell* colIter = rowIter;

            while (colIter != NULL) {

                if (colIter->type==INT_TYPE) {
                    wrt << colIter->data.intData;
                }
                else if (colIter->type == FLOAT_TYPE) {
                    wrt << colIter->data.floatData;
                }
                else if (colIter->type == STRING_TYPE) {
                    wrt << colIter->data.strData;
                }
                else {
                    wrt << "NONE";
                }
                if (colIter->right != NULL) {
                    wrt << ",";
                }
                colIter = colIter->right;
            }
            wrt << "\n";
            rowIter = rowIter->down;
        }
        wrt.close();
        cout << "Sheet Saved Successfully To " << fileName;
    }
    bool IsInteger(const  string& s) {
        try {
             stoi(s);  
            return true;
        }
        catch (...) {
            return false;
        }
    }
    bool IsFloat(const  string& s) {
        try {
             stof(s);  
            return true;
        }
        catch (...) {
            return false;
        }
    }
    void LoadSheet(const  string& fileName) {
         ifstream inFile(fileName);
        if (!inFile) {
             cout << "Error opening file for loading!" <<  endl;
            return;
        }
         vector< vector< string>> fileData;
         string line;

        while (getline(inFile, line)) {
             istringstream iss(line);
             string value;
             vector< string> rowData;

            while (getline(iss, value, ',')) {
                rowData.push_back(value);
            }

            fileData.push_back(rowData);
        }
        inFile.close();

        int numRows = fileData.size();
        int numCols = 0;

        if (numRows > 0) {
            numCols = fileData[0].size();
        }

        if (numRows == 0 || numCols == 0) {
             cout << "No data found in the file!" <<  endl;
            return;
        }

        Cell* prevRow = nullptr;
        for (int i = 0; i < numRows;  i++) {
            Cell* newRow = nullptr;
            Cell* prevCol = nullptr;

            for (int j = 0; j < numCols; j++) {
                Cell* newCell = new Cell(); 
                 string value = fileData[i][j];

                if (value == "NONE") {
                    newCell->type = NONE;
                }
                else if (IsInteger(value)) {
                    newCell->type = INT_TYPE;
                    newCell->data.intData =  stoi(value);
                }
                else if (IsFloat(value)) {
                    newCell->type = FLOAT_TYPE;
                    newCell->data.floatData =  stof(value);
                }
                else {
                    newCell->type = STRING_TYPE;
                    newCell->data.strData = value;
                }

                if (prevCol != nullptr) {
                    prevCol->right = newCell;
                    newCell->left = prevCol;
                }

                prevCol = newCell;

                if (i == 0 && j == 0) {
                    root = newCell;
                }

                if (prevRow != nullptr) {
                    Cell* upperCell = prevRow;
                    for (int k = 0; k < j; ++k) {
                        upperCell = upperCell->right;
                    }
                    upperCell->down = newCell;
                    newCell->up = upperCell;
                }

                if (j == 0) {
                    newRow = newCell;
                }
            }

            prevRow = newRow;
        }

        current = root;

         cout << "Sheet loaded successfully from " << fileName << "!" <<  endl;
    }
    void HandleSaveSheet() {
         string fileName;
         cout << "Do you want to save this sheet? (yes/no): ";
         string response;
         cin >> response;
        if (response == "yes") {
             cout << "Enter a name for the sheet: ";
             cin >> fileName;
            saveSheet(fileName + ".txt");
        }
    }
    void HandleLoadSheet() {
         string fileName;
         cout << "Enter the name of the sheet to load: ";
         cin >> fileName;
        LoadSheet(fileName + ".txt");
    }
    void moveByArrowKeys() {
        while (true) {
            displayGrid();
            setConsoleColor(GREEN);
             cout << "\n\nUse arrow keys to move (or press 'x' to quit).\n";


            int ch = _getch();

          
            if (ch == 13) {
                _getch();
                setCurrentData();
            }

            if (_kbhit()) {
                int ch2 = _getch(); 

                if (ch == 224) {  
                    if (ch2 == 72) { 
                        moveCurrent('U');
                    }
                    else if (ch2 == 80) { 
                        moveCurrent('D');
                    }
                    else if (ch2 == 75) { 
                        moveCurrent('L');
                    }
                    else if (ch2 == 77) { 
                        moveCurrent('R');
                    }
                    Sleep(100);
                }
            }

            if (ch == 49) {
                ClearRow();
            }
            if(ch == 50) {
                ClearCols();
            }
            if (ch == 97) {
                try {
                    deleteRow();
                }
                catch (exception& e) {
                    system("cls");
                    cout<<e.what();
                    return;
                }
            }

            if (ch == 122) {
                try {
                    deleteColumn();
                }
                catch (exception& e) {
                    system("cls");
                    cout << e.what();
                    return;
                }
            }
            if (ch == 116) {
                addRowAbove();
            }
            if (ch == 103) {
                addColRight();
            }
            if (ch == 102) {
                addColLeft();
            }
            if (ch == 118) {
                addRowBelow();
            }
            if (ch == 119) {
                ClearRow();
            }
            if (ch == 101) {
                ClearCols();
            }
            if (ch == 111) {
                InsertCellByRightShift();
            }
            if (ch == 112) {
                InsertCellByDownShift();
            }
            if (ch == 107) {
                DeleteCellByLeftShift();
            }
            if (ch == 108) {
                deleteCellByUpShift();
            }
            if (ch == 114) { 
                testRangeSum();
                _getch();
            }
            if (ch == 117) {
                testRangeAvg();
                _getch();
            }
            if (ch == 113) {
                testMaxMin();
                _getch();
            }
            if (ch == 119) {
                testCountCells();
                _getch();
            }
            if (ch == 98) {
                testCopyCut();
            }
            if (ch == 110) {
                testPaste();
            }
            if (ch == 53) {
                HandleSaveSheet();
                _getch();
            }
            if (ch == 54) {
                HandleLoadSheet();
                _getch();
                system("cls");
                moveByArrowKeys();
            }
            if (ch == 'M' || ch == 'm') {
                displayCommandShortcuts();  
            }
            if (ch == 'x') {
                 cout << "Exiting...\n";
                break;
            }
            system("CLS");
            setConsoleColor(7);
        }
    }
    void drawCell() {
        char ch = -37;
        cout << " ----";
        cout << "\n|";
        cout << "    ";
        cout <<"|\n";
        cout << " ----";
    }
    class excelIT {
    public:
        Cell* currCell;
        excelIT(Cell* currCell) :currCell{ currCell } {}
        Cell& operator*() {
            return *currCell;
        }
        excelIT& operator++() {
            if (currCell->down != NULL) {
                currCell = currCell->down;
            }
            else {
                cout << "Reached The Bottom Boundary...\n";
            }
            return *this;
        }
        excelIT& operator--() {
            if (currCell->up != NULL) {
                currCell = currCell->up;
            }
            else {
                cout << "Reached The Upper Boundary...\n";
            }
            return *this;
        }
        excelIT operator++(int) {
            excelIT temp = *this;
            if (currCell->right != NULL) {
                currCell = currCell->right;
            }
            else {
                cout << "Reached The Right Boundary...\n";
            }
            return temp;
        }
        excelIT operator--(int) {
            excelIT temp = *this;
            if (currCell->left != NULL) {
                currCell = currCell->left;
            }
            else {
                cout << "Reached The Left Boundary...\n";
            }
            return temp;
        }
        bool operator==(const excelIT& other) {
            return currCell == other.currCell;
        }
        bool operator!=(const excelIT& other) {
            return !(currCell == other.currCell);
        }
        Cell* getCurrCell() {
            return currCell;
        }
    };
    excelIT begin() {
        return excelIT(root);
    }
    excelIT end() {
        return excelIT(NULL);
    }
    void iterateRowIncrement() {
        excelIT it = begin();
        while (it.getCurrCell() != NULL) {
            Cell& cell = *it;
            if (cell.type == INT_TYPE) {
                cout << "Cell data: " << cell.data.intData <<"\n";
            }
            if (cell.type == FLOAT_TYPE) {
                cout << "Cell data: " << cell.data.floatData <<"\n";
            }
            if (cell.type == STRING_TYPE) {
                cout << "Cell data: " << cell.data.strData <<"\n";
            }
            if (it.getCurrCell()->right != NULL)
                it++;
            else
                break;
        }
    }
    void iterateRowDecrement() {
        excelIT it = begin();
        while (it.getCurrCell()->right != NULL) {
            it++;
        }
        while (it.getCurrCell() != NULL) {
            Cell& cell = *it;
            if (cell.type == INT_TYPE) {
                cout << "Cell data: " << cell.data.intData <<"\n";
            }
            if (cell.type == FLOAT_TYPE) {
                cout << "Cell data: " << cell.data.floatData <<"\n";
            }
            if (cell.type == STRING_TYPE) {
                cout << "Cell data: " << cell.data.strData <<"\n";
            }
            if (it.getCurrCell()->left != NULL)
                it--;
            else
                break;
        }
    }
    void iterateColumnIncrement() {
        excelIT it = begin();
        while (it.getCurrCell() != NULL) {
            Cell& cell = *it;
            if (cell.type == INT_TYPE) {
                cout << "Cell data: " << cell.data.intData <<"\n";
            }
            if (cell.type == FLOAT_TYPE) {
                cout << "Cell data: " << cell.data.floatData <<"\n";
            }
            if (cell.type == STRING_TYPE) {
                cout << "Cell data: " << cell.data.strData <<"\n";
            }
            if (it.getCurrCell()->down != NULL)
                 ++it;
            else
                break;
        }
    }
    void iterateColumnDecrement() {
        excelIT it = begin();
        while (it.getCurrCell()->down != NULL) {
             ++it;
        }
        while (it.getCurrCell() != NULL) {
            Cell& cell = *it;
            if (cell.type == INT_TYPE) {
                cout << "Cell data: " << cell.data.intData <<"\n";
            }
            if (cell.type == FLOAT_TYPE) {
                cout << "Cell data: " << cell.data.floatData <<"\n";
            }
            if (cell.type == STRING_TYPE) {
                cout << "Cell data: " << cell.data.strData <<"\n";
            }
            if (it.getCurrCell()->up != NULL)
                --it;
            else
                break;
        }
    }
};
void setConsoleColor(int color) {
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), color);
}
void test101() {
    Excel excelSheet;
    excelSheet.moveByArrowKeys();

    system("cls");
    cout << "\t\t\tORIGINAL GRID LATEST STATE: \n\n";
    excelSheet.displayGrid();

    cout << "\t\t\tPRINTING THE DATA / CONTENTS OF GRID THROUGH ITERATORS\n\n";

    cout << "Iterating over rows (Increment):" <<"\n";
    excelSheet.iterateRowIncrement();

    cout << "\nIterating over rows (Decrement):" <<"\n";
    excelSheet.iterateRowDecrement();

    cout << "\nIterating over columns (Increment):" <<"\n";
    excelSheet.iterateColumnIncrement();

    cout << "\nIterating over columns (Decrement):" <<"\n";
    excelSheet.iterateColumnDecrement();

    setConsoleColor(7);

}
void displayHomePage() {

    setConsoleColor(GREEN);
    cout << "  _    _ _   _______ _____ __  __       _______ ______   ________   _______ ______ _      \n";
    cout << " | |  | | | |__   __|_   _|  \\/  |   /\\|__   __|  ____| |  ____\\ \\ / / ____|  ____| |     \n";
    cout << " | |  | | |    | |    | | | \\  / |  /  \\  | |  | |__    | |__   \\ V / |    | |__  | |     \n";
    cout << " | |  | | |    | |    | | | |\\/| | / /\\ \\ | |  |  __|   |  __|   > <| |    |  __| | |     \n";
    cout << " | |__| | |____| |   _| |_| |  | |/ ____ \\| |  | |____  | |____ / . \\ |____| |____| |____ \n";
    cout << "  \\____/|______|_|  |_____|_|  |_/_/    \\_\\_|  |______| |______/_/ \\_\\_____|______|______|\n\n";

    cout << " **********************************************\n";
    cout << " *         Made By:                           *\n";
    cout << " *         Mohid Faisal BSCS23092             *\n";
    cout << " *                                            *\n";
    cout << " **********************************************\n";
    setConsoleColor(7);

}
int main() {
    displayHomePage();
    _getch();
    system("cls");
    test101();
    return 0;
}

//t add row up
//v add row down
//g add col right
//f add col left
//a del row
//z del col
//o insert cell right shift
//p insert cell down shift
//k deletebyleftshift
//l deletebyupshift
//r range sum
//u range avg
//w range count
//q range max-min
//b copy-cut
//n paste
//arrows to move inside the grid