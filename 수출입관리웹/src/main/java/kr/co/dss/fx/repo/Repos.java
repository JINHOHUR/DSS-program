package kr.co.dss.fx.repo;

import kr.co.dss.fx.domain.*;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;
import java.util.Optional;

/** 저장소 모음. 작아서 한 파일에 둔다. */
public final class Repos {
    private Repos() {}

    public interface DealRepo extends JpaRepository<Deal, Long> {
        @Query("select d from Deal d where d.deletedAt is null order by d.docNo desc")
        List<Deal> findAllActive();

        @Query("select d.docNo from Deal d where d.docNo like ?1")
        List<String> docNosLike(String prefix);
    }

    public interface StageLogRepo extends JpaRepository<StageLog, Long> {
        List<StageLog> findByDealId(Long dealId);
        Optional<StageLog> findByDealIdAndStageNo(Long dealId, int stageNo);
        void deleteByDealId(Long dealId);
    }

    public interface CompanyRepo extends JpaRepository<Company, String> {
        List<Company> findByKindOrderBySortAscNameAsc(String kind);
        List<Company> findAllByOrderByKindAscSortAscNameAsc();
    }

    public interface CodeRepo extends JpaRepository<CodeItem, Long> {
        List<CodeItem> findByCategoryOrderBySortAscValueAsc(String category);
        Optional<CodeItem> findByCategoryAndValue(String category, String value);
    }

    public interface SettingRepo extends JpaRepository<Setting, String> {}

    public interface LcRuleRepo extends JpaRepository<LcRule, String> {}

    public interface FlowNodeRepo extends JpaRepository<FlowNode, Long> {
        List<FlowNode> findAllByOrderBySeqAscIdAsc();
    }

    public interface FlowEdgeRepo extends JpaRepository<FlowEdge, Long> {
        List<FlowEdge> findAllByOrderByIdAsc();
        Optional<FlowEdge> findBySrcAndDst(Long src, Long dst);
        void deleteBySrcOrDst(Long src, Long dst);
    }
}
