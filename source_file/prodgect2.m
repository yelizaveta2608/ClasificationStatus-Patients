function varargout = prodgect2(varargin)
% PRODGECT2 MATLAB code for prodgect2.fig
%      PRODGECT2, by itself, creates a new PRODGECT2 or raises the existing
%      singleton*.
%
%      H = PRODGECT2 returns the handle to a new PRODGECT2 or the handle to
%      the existing singleton*.
%
%      PRODGECT2('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in PRODGECT2.M with the given input arguments.
%
%      PRODGECT2('Property','Value',...) creates a new PRODGECT2 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before prodgect2_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to prodgect2_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help prodgect2

% Last Modified by GUIDE v2.5 28-Jan-2019 13:26:58

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @prodgect2_OpeningFcn, ...
                   'gui_OutputFcn',  @prodgect2_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before prodgect2 is made visible.
function prodgect2_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to prodgect2 (see VARARGIN)

% Choose default command line output for prodgect2
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes prodgect2 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = prodgect2_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



function edit1_Callback(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit1 as text
%        str2double(get(hObject,'String')) returns contents of edit1 as a double


% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit2_Callback(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit2 as text
%        str2double(get(hObject,'String')) returns contents of edit2 as a double


% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit3_Callback(hObject, eventdata, handles)
% hObject    handle to edit3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit3 as text
%        str2double(get(hObject,'String')) returns contents of edit3 as a double


% --- Executes during object creation, after setting all properties.
function edit3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
close prodgect2


function edit4_Callback(hObject, eventdata, handles)
% hObject    handle to edit4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit4 as text
%        str2double(get(hObject,'String')) returns contents of edit4 as a double


% --- Executes during object creation, after setting all properties.
function edit4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit5_Callback(hObject, eventdata, handles)
% hObject    handle to edit5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit5 as text
%        str2double(get(hObject,'String')) returns contents of edit5 as a double


% --- Executes during object creation, after setting all properties.
function edit5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit6_Callback(hObject, eventdata, handles)
% hObject    handle to edit6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit6 as text
%        str2double(get(hObject,'String')) returns contents of edit6 as a double


% --- Executes during object creation, after setting all properties.
function edit6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit7_Callback(hObject, eventdata, handles)
% hObject    handle to edit7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit7 as text
%        str2double(get(hObject,'String')) returns contents of edit7 as a double


% --- Executes during object creation, after setting all properties.
function edit7_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global op kko dcc kkc nkkc kke de rhnV dcS dcH dcV kkd kdtc kkt kkn kne nomer nam u u1 u2 u3 u4 u5 u6 u7 u8 u9 u10 kkz u11;
global osC osD osO osT osV osN kvN osE kvE osOL
set(handles.edit1, 'String',['Коефіцієнт = ' num2str(dcc), ' ' kkc])
set(handles.edit4, 'String',['К =' num2str(dcS),';',num2str(dcH),';',num2str(dcV),'-' kkd])
set(handles.edit5, 'String',['Коефіцієнт = ' num2str(op), ' ' kko])
set(handles.edit6, 'String',['Коефіцієнт =' num2str(kdtc), ' ' kkt])
set(handles.edit7, 'String',['Коефіцієнт = ' num2str(rhnV), ' ' kne])
set(handles.edit3, 'String',['Коефіцієнт = ' num2str(nkkc), ' ' kkn])
set(handles.edit2, 'String',['Коефіцієнт = ' num2str(de), ' ' kke])
osOL=(0.225*osE+0.225*osN+0.133*osD+0.116*osC+0.079*osT+0.075*osV+0.072*osO+0.0375*kvE+0.0375*kvN)*100

% nam=4;
nam=xlsread('zhurnal.xls','A', 'AO1');
u=num2str(nam);
     
     u1=['A'];  u1=[u1 u]; 
     u2=['B'];  u2=[u2 u]; 
     u3=['C'];  u3=[u3 u];
     u4=['D'];  u4=[u4 u]; 
     u5=['E'];  u5=[u5 u]; 
     u6=['F'];  u6=[u6 u];
     u7=['G'];  u7=[u7 u];
     u8=['H'];  u8=[u8 u];
     u9=['I'];  u9=[u9 u];
     u10=['J'];  u10=[u10 u];
     u11=['K'];  u11=[u11 u];
     xlswrite('zhurnal.xls',nomer,'A', u1);
     xlswrite('zhurnal.xls',dcc,'A', u2);
     xlswrite('zhurnal.xls',dcS,'A', u3);
     xlswrite('zhurnal.xls',dcH,'A', u4);
     xlswrite('zhurnal.xls',dcV,'A', u5);
     xlswrite('zhurnal.xls',op,'A', u6);
     xlswrite('zhurnal.xls',kdtc,'A', u7);
     xlswrite('zhurnal.xls',rhnV,'A', u8);
     xlswrite('zhurnal.xls',nkkc,'A', u9);
     xlswrite('zhurnal.xls',de,'A', u10);
     nam=nam+1;
     xlswrite('zhurnal.xls',nam,'A', 'AO1');
     if dcc<2.1 && 1.2<dcV && dcV<2.6 && op>8 && op<10 && kdtc<0.1 && rhnV<10 && nkkc<3.3 && de==1
         kkz='Пацієнт здоровий. Організм в функціональному стані.';
     elseif dcc>2.1 && dcV>2.6 && op>10 && kdtc>0.1 && rhnV>10 && de~=1 && nkkc>6
         kkz='Пацієнт хворий. Організм в не функціональному стані.';
     elseif dcc>2.1 && 1.2>dcV && op<8 && kdtc>0.1 && rhnV>10 && de~=1&& nkkc>6
         kkz='Пацієнт хворий. Організм в не функціональному стані.';
     elseif dcc>2.1 | dcV<2.6 | op>10 | kdtc>0.1 | rhnV>10 | de~=1 | 1.2>dcV | op<8 | nkkc>6
         kkz='Одна з систем не в функціональному стані';
     else   
         kkz='Більше одної системи погано функціонує';
     end
set(handles.edit8, 'String',['k=' num2str(osOL) '% ' kkz ])
xlswrite('zhurnal.xls',{kkz},'A', u11);
clear dcc dcV op kdtc rhnV de nkkc kkc kkd kko kkt kkv kkn kke osOL kkz 



function edit8_Callback(hObject, eventdata, handles)
% hObject    handle to edit8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit8 as text
%        str2double(get(hObject,'String')) returns contents of edit8 as a double


% --- Executes during object creation, after setting all properties.
function edit8_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
