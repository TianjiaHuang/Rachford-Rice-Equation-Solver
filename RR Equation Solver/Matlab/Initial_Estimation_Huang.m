function centroid = Initial_Estimation_Huang(z, K)
%INITIAL_ESTIMATION Compute the centroid of the smaller feasible region.
%
% This function constructs the feasible region defined by the inequality
% constraints associated with z and K, computes all valid intersection
% points of the constraint boundaries, removes duplicate vertices, and
% returns the centroid of the feasible region vertices.
%
% Inputs:
%   z : 1 x NC vector of overall compositions
%   K : (NP-1) x NC matrix of K-values
%
% Output:
%   centroid : (NP-1) x 1 centroid of the feasible-region vertices
%
% Notes:
%   - The feasible region is defined by A * x <= b.
%   - Each vertex is obtained from the intersection of NP-1 active
%     constraint planes.
%   - If no feasible region is found, centroid is returned as NAN.

    [NPm1, NC] = size(K);   % NPm1 = NP - 1

    % Total number of inequality constraints:
    %   NC composition-related constraints
    %   NPm1 non-negativity constraints: beta > 0
    %   1 upper-sum constraint: sum(beta) < 1
    %   Please notice the condition below
    %   If positive flash: 
    numConstraints = NC + NPm1 + 1; 
    %   If negative flash (beta can < 0)
    % numConstraints = NC; 

    A = zeros(numConstraints, NPm1);
    b = zeros(numConstraints, 1);

    %--------------------------------------------------------------
    % Step 1: Build the composition-related inequality constraints
    %--------------------------------------------------------------
    for i = 1:NC
        constraintVector = 1 - K(:, i);
        theta = 1;
        for j = 1: NPm1
            [K1, ~] = max(K(j, :));
            [Kn, ~] = min(K(j, :));
            if K(j,i)>1
                theta = min(theta,(1-Kn)/(K(j,i)-Kn));
            else
                theta = min(theta,(K1-1)/(K1-K(j,i)));
            end
        end
        constraintBound = min([1 - z(i)/theta, min(1 - K(:, i) * z(i))]);

        A(i, :) = constraintVector';
        b(i) = constraintBound;
    end

    %--------------------------------------------------------------
    % Step 2: Add additional constraints for positive flash
    %   beta_j >= 0    ->   -beta_j <= 0
    %   sum(beta_j) <= 1
    %--------------------------------------------------------------
    if numConstraints > NC
        for j = 1:NPm1
            A(NC + j, j) = -1;
            b(NC + j) = 0;
        end
    
        A(numConstraints, :) = 1;
        b(numConstraints) = 1;
    end

    %--------------------------------------------------------------
    % Step 3: Compute all candidate vertices by intersecting
    %         NPm1 constraint hyperplanes at a time
    %--------------------------------------------------------------
    vertices = [];
    activeSets = nchoosek(1:numConstraints, NPm1);

    for i = 1:size(activeSets, 1)
        selectedRows = activeSets(i, :);
        M = A(selectedRows, :);
        rhs = b(selectedRows);

        % A valid vertex requires a full-rank system
        if rank(M) == NPm1
            point = M \ rhs;

            % Keep the point only if it satisfies all inequalities
            if all(A * point <= b + 1e-8)
                vertices(end+1, :) = point';
            end
        end
    end

    %--------------------------------------------------------------
    % Step 4: Remove duplicate vertices
    %--------------------------------------------------------------
    if ~isempty(vertices)
        vertices = unique(round(vertices, 8), 'rows');
    end

    %--------------------------------------------------------------
    % Step 5: Compute the centroid of the feasible-region vertices
    %--------------------------------------------------------------
    if isempty(vertices)
        warning('No feasible region found.');
        centroid = NaN(NPm1, 1);
    else
        centroid = mean(vertices, 1)';
    end
end