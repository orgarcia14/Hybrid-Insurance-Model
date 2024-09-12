function Incident = states(Ii,Delta,Theta)
Incident = [];
for i = 1:5
    I_i = [];
    for j = 1:length(Ii)
        I_i(j) = (1-Theta(i)*Delta(j))*Ii(j);
    end
    I_i = transpose(I_i);
    Incident = [Incident I_i];
end
