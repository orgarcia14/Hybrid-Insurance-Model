clc
clear all

Zn = readmatrix("CI_data.xlsx","Sheet","Results","Range","B5:D14");
relative_n = readmatrix("CI_data.xlsx","Sheet","Results","Range","E5:E14");

Zx = readmatrix("CI_data.xlsx","Sheet","Results","Range","H5:J40");
relative_x = readmatrix("CI_data.xlsx","Sheet","Results","Range","K5:K40");

n = Zn(:,1);
new_n = Zn(:,2);
ci_n = Zn(:,3);

x = Zx(:,1);
new_x = Zx(:,2);
ci_x = Zx(:,3);

% Plot for N
figure()
plot(n,new_n,"--ob","MarkerFaceColor","b","MarkerSize",3)
hold on
plot(n,ci_n,"-*g","MarkerFaceColor","g")
grid on
lgd=legend("New Model",'CI Model','Location','north');

xlabel('($n$) term','FontSize',12,'Interpreter','latex')
ylabel('Mean of Benefit','FontSize',12)

% Plot for relative diff n
figure()
plot(n,relative_n,"-k")
hold on
plot(n,relative_n-relative_n,"--b")
xlabel('($n$) term','FontSize',12,'Interpreter','latex')
ylabel('Relative difference of the means','FontSize',12)


%Plot for X

figure()
plot(x,new_x,"--b")
hold on
plot(x,ci_x,"-g")
grid on
lgd=legend("New Model",'CI Model','Location','north');

xlabel('($x$) age','FontSize',12,'Interpreter','latex')
ylabel('Mean of Benefit','FontSize',12)

% Plot for relative diff x
figure()
plot(x,relative_x,"-k")
hold on
plot(x,relative_x-relative_x,"--b")
xlabel('($x$) age','FontSize',12,'Interpreter','latex')
ylabel('Relative difference of the means','FontSize',12)


%% Reserves

statereserve = readmatrix("CI_data.xlsx","Sheet","Results","Range","AB5:AG25");
averagereserve = readmatrix("CI_data.xlsx","Sheet","Results","Range","AK5:AL25");

t = statereserve(:,1);
s1 = statereserve(:,2);
s2 = statereserve(:,3);
s3 = statereserve(:,4);
s4 = statereserve(:,5);
s5 = statereserve(:,6);

reserve_new = averagereserve(:,1);
reserve_ci = averagereserve(:,2);

%Plot for reserve per state

figure()
plot(t,s1,"--","Color","#A2142F")
hold on
plot(t,s2,"--","Color","#7E2F8E")
plot(t,s3,"--","Color","#0072BD")
plot(t,s4,"--","Color","#77AC30")
plot(t,s5,"--","Color","#EDB120")
plot(t,reserve_ci, "-k")
grid on
lgd=legend("State 1",'State 2','State 3','State 4','State 5','CI model','Location','northwest');
xlabel('policy year $t$','FontSize',12,'Interpreter','latex')
ylabel('Reserve','FontSize',12)

%Plot for relative diff
figure()
plot(t,s1-reserve_ci,"--","Color","#A2142F")
hold on
plot(t,s2-reserve_ci,"--","Color","#7E2F8E")
plot(t,s3-reserve_ci,"--","Color","#0072BD")
plot(t,s4-reserve_ci,"--","Color","#77AC30")
plot(t,s5-reserve_ci,"--","Color","#EDB120")
plot(t,reserve_ci-reserve_ci,"-k")
grid on
lgd=legend("State 1",'State 2','State 3','State 4','State 5','CI model','Location','northeast');
xlabel('policy year $t$','FontSize',12,'Interpreter','latex')
ylabel('Relative difference of the reserve','FontSize',12)


%Plot of average reserve
figure()
plot(t,reserve_new,"--b")
hold on
plot(t,reserve_ci,"-g")
lgd=legend("New Model",'CI Model','Location','northeast');
xlabel('policy year $t$','FontSize',12,'Interpreter','latex')
ylabel('Reserve','FontSize',12)

figure()
plot(t,reserve_ci-reserve_new,"-k")
hold on
plot(t,reserve_ci-reserve_ci,"--b")
xlabel('policy year $t$','FontSize',12,'Interpreter','latex')
ylabel('Reserve','FontSize',12)



%% Premiums

year = readmatrix("CI_data.xlsx","Sheet","Results","Range","Q9:Q28");
premium = readmatrix("CI_data.xlsx","Sheet","Results","Range","W9:W28");

ci_prem = zeros(length(year),1);
ci_prem = ci_prem + 2460.44;

figure()
plot(year,premium,"--ob")
hold on
plot(year,ci_prem,"-*g")
lgd=legend("New Model (Average)",'CI Model','Location','north');
xlabel('policy year $t$','FontSize',12,'Interpreter','latex')
ylabel('Premium','FontSize',12)


