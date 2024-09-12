clc
clear all

% Read data in Excel file
Ic = readmatrix("CI_data.xlsx","Sheet","IncidentRate","Range","B3:B72");       % Cancer Incident
Ih = readmatrix("CI_data.xlsx","Sheet","IncidentRate","Range","C3:C72");       % Heart disease Incident
Is = readmatrix("CI_data.xlsx","Sheet","IncidentRate","Range","D3:D72");       % Stroke Incident
Id = readmatrix("CI_data.xlsx","Sheet","IncidentRate","Range","E3:E72");       % Death due to Other causes

DeltaC = readmatrix("CI_data.xlsx","Sheet","AttribFrac","Range","C3:C72");     % Attributable fraction for Cancer
DeltaH = readmatrix("CI_data.xlsx","Sheet","AttribFrac","Range","F3:F72");     % Attributable fraction for Heart Disease
DeltaS = readmatrix("CI_data.xlsx","Sheet","AttribFrac","Range","I3:I72");     % Attributable fraction for Stroke
DeltaO = readmatrix("CI_data.xlsx","Sheet","AttribFrac","Range","L3:L72");     % Attributable fraction for other causes

Tau = readmatrix("CI_data.xlsx","Sheet","TransitionProb","Range","D5:H9");     % Transition Probability



% Other Data
Theta = [-0.5; -0.1; 0.3; 0.75; 1.0];       %Percentage of removed attributable mortality
omega = [0.30; 0.35; 0.20; 0.10; 0.05];     % Initial Distribution
r = [0 0.05 0.1 0.15 0.2];                  % Reduction Rate

age = 25;       % Age
x = age-19;

n = 20;         % Term duration
v = 1/1.06;     % Interest Rate

% Benefits
DB = 1000000;    % Death due to other causes
CB = 1000000;    % Cancer
HB = 1000000;    % Heart Disease
SB = 1000000;    % Stroke

% Create incident data for each states
IC = transpose(states(Ic,DeltaC,Theta));
IH = transpose(states(Ih,DeltaH,Theta));
IS = transpose(states(Is,DeltaS,Theta));
ID = transpose(states(Id,DeltaO,Theta));

Qnew = IC + IH + IS + ID;
m = width(Qnew);

% Create data for Phi and P_x
Phi = {};    % Create cell array for phi
P = {};      % Create cell array for P_x

for i=1:m
    Phix = 1 - Qnew(:,i);
    Phi{i} = diag(Phix);
    P{i} = Phi{i}*Tau;
end

% Calculation for PV Benefit of Critical Insurance
    % Cancer
    ZC = 0;
    for s=1:5
        BsC = IBenefit(x,n,v,s,P,IC,1);
        ZC = ZC + omega(s).*BsC;
    end
    ZC = CB*ZC;

    % Heart Attack
    ZH = 0;
    for s=1:5
        BsH = IBenefit(x,n,v,s,P,IH,1);
        ZH = ZH + omega(s).*BsH;
    end
    ZH = HB*ZH;

    % Stroke
    ZS = 0;
    for s=1:5
        BsS = IBenefit(x,n,v,s,P,IS,1);
        ZS = ZS + omega(s).*BsS;
    end
    ZS = SB*ZS;

    % Death due to other causes
    ZD = 0;
    for s=1:5
        BsD = IBenefit(x,n,v,s,P,ID,1);
        ZD = ZD + omega(s).*BsD;
    end
    ZD = DB*ZD;

    ZQ = 0;
    for s=1:5
        BsD = IBenefit(x,n,v,s,P,Qnew,1);
         ZQ = ZQ + omega(s).*BsD;
    end
    ZQ = DB*ZQ;


Z = ZC + ZH + ZS + ZD; % Total PV of Benefits
display(Z)
display(ZQ)


% Calculation for PV of Annuity
C = 0;
for s=1:5
    as = annuity(x,n,v,s,P,r);
    C = C + omega(s).*as;
end


% Premium 
pi = (Z)/C;
display(pi)

writematrix(pi,"CI_data.xlsx","Sheet","Results","Range", "N6",'AutoFitWidth',0)


% Reserves
V = zeros(n+1, length(omega));
for s=1:length(omega)
    for k=0:n
        BsC = IBenefit(x+k,n-k,v,s,P,IC,1);
        BsH = IBenefit(x+k,n-k,v,s,P,IH,1);
        BsS = IBenefit(x+k,n-k,v,s,P,IS,1);
        BsD = IBenefit(x+k,n-k,v,s,P,ID,1);
        as = annuity(x+k,n-k,v,s,P,r);
        V(k+1,s) = (CB*BsC + HB*BsH + SB*BsS + DB*BsD) - pi.*as;
    end
end
writematrix(V,"CI_data.xlsx","Sheet","Results","Range", "AC5",'AutoFitWidth',0,'PreserveFormat',true)



% Status Distribution through Px
sn=30;
distri = zeros(sn,5);
for t=0:sn-1
    tpx = survivalP(t,x,P);
    for l=1:5
        distri(t+1,l) = (transpose(omega)*tpx(:,l));
    end
    distri(t+1,:) = distri(t+1,:)/sum(distri(t+1,:));
end
writematrix(distri,"CI_data.xlsx","Sheet","Status Distribution","Range", "K5",'AutoFitWidth',0,'PreserveFormat',true)

% Status Distribution through P
Distri = zeros(sn, length(omega));
for i=1:sn
    Distri(i,:) = transpose(omega)*(Tau^(i-1));
end
writematrix(Distri,"CI_data.xlsx","Sheet","Status Distribution","Range", "C5",'AutoFitWidth',0,'PreserveFormat',true)
%% Mean of Benefit for a 25 yr old with n-year temp insurance 

ni = 5:5:50;
xi = 25-19;

Zi = zeros(length(ni),1);
for i=1:length(ni)
    Zis = 0;
    for s=1:5
        BsD = IBenefit(xi,ni(i),v,s,P,Qnew,1);
        Zis = Zis + omega(s).*BsD;
    end
    Zi(i) = DB*Zis;
end

% Write the data into Excel file
writematrix(transpose(ni),"CI_data.xlsx","Sheet","Results","Range", "B5",'AutoFitWidth',0,'PreserveFormat',true) % Term
writematrix(Zi,"CI_data.xlsx","Sheet","Results","Range", "C5",'AutoFitWidth',0,'PreserveFormat',true)            % New Model

%% Mean of Benefit for a x-yr old with 20-year temp insurance

nj = 20;
agej = 25:1:60;
xj = agej - 19;

Zj = zeros(length(xj),1);
for j=1:length(xj)
    Zjs = 0;
    for s=1:5
        BsD = IBenefit(xj(j),nj,v,s,P,Qnew,1);
        Zjs = Zjs + omega(s).*BsD;
    end
    Zj(j) = DB*Zjs;
end

% Write the data into Excel file
writematrix(transpose(agej),"CI_data.xlsx","Sheet","Results","Range", "H5",'AutoFitWidth',0,'PreserveFormat',true)  % Age
writematrix(Zj,"CI_data.xlsx","Sheet","Results","Range", "I5",'AutoFitWidth',0,'PreserveFormat',true)               % New Model


